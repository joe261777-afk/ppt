"""Streamlit entry point for the workshop evaluation generator."""
from pathlib import Path
import tempfile
import traceback
import pandas as pd
import streamlit as st
import report_engine as engine
from openpyxl import load_workbook

st.set_page_config(page_title='Workshop Evaluation Generator', page_icon='📊', layout='wide')
st.title('Workshop Evaluation Report Generator')
st.caption('Upload workshop information, verify attendance and download the finished PowerPoint. Both templates are included.')

@st.cache_resource
def templates():
    return engine.install_embedded_templates()

@st.cache_data(show_spinner=False)
def uploaded_data(file_bytes, suffix):
    # Current file bytes are securely kept in the Streamlit session cache.
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        return tmp.name

def save_upload(file):
    return uploaded_data(file.getvalue(), Path(file.name).suffix.lower())

try:
    report_template, attendance_template = templates()
except Exception as e:
    st.error('PowerPoint templates are missing. Upload the included templates/ folder to your GitHub repository.')
    st.exception(e)
    st.stop()

left,right=st.columns(2)
with left:
    survey_pdf=st.file_uploader('1 · Survey results PDF',type=['pdf'])
with right:
    response_xlsx=st.file_uploader('2 · Survey responses Excel',type=['xlsx','xlsm'])

mode=st.radio('3 · Attendance source', ['Excel sheet','Scanned image'],horizontal=True)
attend_input=st.file_uploader('Attendance Excel file' if mode=='Excel sheet' else 'Attendance photograph', type=['xlsx','xlsm'] if mode=='Excel sheet' else ['jpg','jpeg','png','webp'],key='attendance_file')

if st.button('Read attendance',type='secondary',disabled=attend_input is None):
    try:
        input_path=save_upload(attend_input)
        if mode=='Excel sheet':
            wb=load_workbook(input_path,data_only=True)
            found=[result for ws in wb.worksheets if (result:=engine.parse_sheet(ws))]
            if not found:raise ValueError('No names and attendance day columns found in the spreadsheet.')
            parsed=max(found,key=lambda p:len(p[0]))
        else:
            parsed=engine.read_attendance_image(input_path)
        people,day_labels,source=parsed
        if len(people)>76:raise ValueError('The original attendance slide supports at most 76 names.')
        st.session_state['attendance_rows']=pd.DataFrame([[name]+statuses+['']*(2-len(statuses)) for name,statuses in people],columns=['Participant name','Day 1','Day 2'])
        st.session_state['days']=day_labels
        st.session_state['review_source']=source
        st.session_state['review_mode']=mode
        st.session_state['review_hash']=hash(attend_input.getvalue())
        st.session_state.pop('result_files',None)
    except Exception as e:
        st.error(f'Could not read attendance: {e}')

if 'attendance_rows' in st.session_state:
    st.subheader('Review and correct attendance')
    st.info('Check every name and mark against the source file, especially for scanned images. Use present, absent or unknown.')
    dcount=len(st.session_state['days'])
    d1=st.text_input('Day 1 label',value=st.session_state['days'][0],key='date_day1')
    d2=st.text_input('Day 2 label',value=st.session_state['days'][1] if dcount>1 else '',key='date_day2',disabled=dcount==1)
    rows=st.data_editor(st.session_state['attendance_rows'],hide_index=True,num_rows='dynamic',use_container_width=True,
        column_config={
            'Participant name':st.column_config.TextColumn('Participant name',required=True),
            'Day 1':st.column_config.SelectboxColumn('Day 1',options=['present','absent','unknown'],required=True),
            'Day 2':st.column_config.SelectboxColumn('Day 2',options=['present','absent','unknown'] if dcount==2 else [''],disabled=dcount==1),
        },key='attendance_editor')
    confirm=st.checkbox('I have reviewed the extracted names and attendance marks against the original image.',value=False) if st.session_state['review_mode']=='Scanned image' else True
    if st.button('Generate PowerPoint',type='primary'):
        if not survey_pdf or not response_xlsx or not attend_input:
            st.error('Please upload the survey PDF, survey responses Excel and attendance input.')
        elif st.session_state['review_mode']!=mode or st.session_state['review_hash']!=hash(attend_input.getvalue()):
            st.error('Attendance source changed. Click Read attendance again.')
        elif not confirm:
            st.error('Please review and confirm the image attendance before generation.')
        else:
            try:
                labels=[d1.strip()] + ([d2.strip()] if dcount==2 else [])
                if not all(labels):raise ValueError('Attendance date label cannot be blank.')
                people=[]
                for idx,row in enumerate(rows.fillna('').itertuples(index=False,name=None),1):
                    name=str(row[0]).strip()
                    if not name or name.startswith('[REVIEW NAME'):
                        raise ValueError(f'Correct the name in row {idx}.')
                    statuses=[str(value).strip().lower() for value in row[1:1+dcount]]
                    if any(status not in ('present','absent','unknown') for status in statuses):
                        raise ValueError(f'Correct the status in row {idx}.')
                    people.append((name,statuses))
                if not 1<=len(people)<=76:raise ValueError('Attendance must contain 1–76 participants.')
                parsed=(people,labels,st.session_state['review_source'])
                job_dir=Path(tempfile.mkdtemp(prefix='evaluation_'))
                result=job_dir/'Evaluation_Report_With_Attendance_FINAL.pptx'
                report=job_dir/'Evaluation_Report_Automated.pptx'
                attend=job_dir/'Attendance_Only_One_Slide_Striped.pptx'
                with st.spinner('Generating your evaluation report...'):
                    engine.build(report_template,save_upload(survey_pdf),save_upload(response_xlsx),str(report),attendance_people=people,attendance_days=dcount)
                    engine.make_attendance_one_slide(save_upload(attend_input),attendance_template,output=str(attend),parsed=parsed)
                    previous=engine.REPORT_PPTX
                    try:
                        engine.REPORT_PPTX=str(report)
                        engine.combine_report_and_attendance(str(attend),output=str(result))
                    finally:
                        engine.REPORT_PPTX=previous
                audit=report.with_suffix('.audit.json')
                st.session_state['result_files']=(result.read_bytes(),audit.read_bytes())
                st.success(f'Report ready: {len(people)} participants, {sum("present" in s for _,s in people)} attended at least once. Attendance is slide 4.')
            except Exception as e:
                st.error(f'Generation failed: {e}')
    if 'result_files' in st.session_state:
        out,audit=st.session_state['result_files']
        st.download_button('Download PowerPoint',out,file_name='Evaluation_Report_With_Attendance_FINAL.pptx',mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',type='primary')
        st.download_button('Download audit JSON',audit,file_name='Evaluation_Report_Automated.audit.json',mime='application/json')

st.caption('Files are processed by the application host. Use Streamlit access controls for confidential participant names, signatures and feedback.')
