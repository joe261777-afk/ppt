"""Streamlit entry point for the workshop evaluation generator."""
from pathlib import Path
import tempfile
import traceback
import pandas as pd
import streamlit as st
import report_engine as engine
from openpyxl import load_workbook
from zoho_library import Library
from sync_queue import new_request, get_request, valid_url

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

st.subheader('Survey data')
source = st.radio('Where should survey data come from?',
                  ['Saved Zoho project (automatic lookup)', 'Upload PDF + Excel manually'], horizontal=True)
project_pdf_bytes = None
project_excel_bytes = None
project_filename = None
survey_pdf = None
response_xlsx = None
if source == 'Saved Zoho project (automatic lookup)':
    try:
        repo = st.secrets.get('PROJECT_LIBRARY_REPO', '')
        token = st.secrets.get('GITHUB_TOKEN', '')
        if not repo or not token:
            st.warning('Project Library is not connected yet. Set PROJECT_LIBRARY_REPO and GITHUB_TOKEN in Streamlit Secrets.')
        else:
            library = Library(repo, token)
            if st.button('Refresh project list') or 'library_projects' not in st.session_state:
                st.session_state['library_projects'] = library.list_projects()
            projects = st.session_state.get('library_projects', [])
            search = st.text_input('Search project name', placeholder='Arab Bank - Flex Your Agility Muscles')
            matches = [p for p in projects if search.casefold() in p['project'].casefold()]
            if matches:
                names = {m['project']:m for m in matches}
                chosen = st.selectbox('Matching surveys', list(names))
                if st.button('Retrieve saved Zoho PDF and spreadsheet'):
                    with st.spinner('Downloading private survey exports...'):
                        pdf_data, xlsx_data = library.load_project(names[chosen])
                    if not pdf_data.startswith(b'%PDF') or xlsx_data[:2]!=b'PK':
                        raise ValueError('Stored files did not match expected PDF and spreadsheet formats.')
                    st.session_state['zoho_files'] = (chosen,pdf_data,xlsx_data)
                if 'zoho_files' in st.session_state:
                    project_filename, project_pdf_bytes, project_excel_bytes = st.session_state['zoho_files']
                    if project_filename != chosen:
                        project_pdf_bytes = project_excel_bytes = None
                    else:
                        st.success(f'Loaded: {project_filename} (PDF + Excel).')
            else:
                st.info('No saved project matches. Use Request fresh Zoho export below to sync it.')
            st.divider()
            st.markdown('#### Request fresh Zoho exports')
            st.caption('This sends a job to the Windows helper. Your Windows computer must be running and signed into Zoho. No Zoho password is sent to Streamlit or GitHub.')
            project_request = st.text_input('Exact Zoho project name', value=search.strip(), key='request_project')
            survey_url = st.text_input('Zoho survey dashboard URL (recommended for accurate matching)', key='request_url', placeholder='https://survey.zoho.com/survey/newui/#/...')
            # Require site authentication; project data is private and sync requests cause remote work.
            sync_password = st.secrets.get('SYNC_BUTTON_PASSWORD', '')
            typed_password = st.text_input('Sync access code', type='password', key='request_password') if sync_password else ''
            if not sync_password:
                st.warning('For security, set SYNC_BUTTON_PASSWORD in Streamlit Secrets to enable remote exports.')
            if st.button('Request fresh Zoho PDF + Excel', disabled=not sync_password):
                if typed_password != sync_password:
                    st.error('Invalid sync access code.')
                else:
                    job = new_request(library, project_request, survey_url)
                    st.session_state['active_sync_id'] = job['id']
                    st.success('Sync request queued. Keep the Windows helper running, then check status.')
            if st.session_state.get('active_sync_id'):
                if st.button('Check Zoho sync status'):
                    job = get_request(library, st.session_state['active_sync_id'])
                    if job:
                        st.write(f"**{job['status'].upper()}** — {job.get('message', '')}")
                        if job['status'] == 'done':
                            st.session_state['library_projects'] = library.list_projects()
                            st.success('Survey updated. Select it above and retrieve the saved files.')
    except Exception as error:
        st.error(f'Project Library lookup failed: {error}')
else:
    left,right=st.columns(2)
    with left:
        survey_pdf=st.file_uploader('1 · Survey results PDF',type=['pdf'])
    with right:
        response_xlsx=st.file_uploader('2 · Survey responses Excel',type=['xlsx','xlsm'])

def selected_survey_files():
    if source == 'Saved Zoho project (automatic lookup)':
        if not (project_pdf_bytes and project_excel_bytes):
            raise ValueError('Choose a saved project and click Retrieve before generating.')
        # Keep the original project name in filename to preserve training metadata parsing.
        with tempfile.NamedTemporaryFile(prefix='zoho_pdf_', suffix='.pdf', delete=False) as fp:
            fp.write(project_pdf_bytes)
            pdf_path=fp.name
        with tempfile.NamedTemporaryFile(prefix='zoho_excel_', suffix='.xlsx', delete=False) as fx:
            fx.write(project_excel_bytes)
            xlsx_path=fx.name
        # The original engine derives metadata from the PDF filename.
        safe_name = ''.join(c if c.isalnum() or c in ' -_()' else '_' for c in project_filename)[:150]
        named_pdf = Path(tempfile.mkdtemp(prefix='zoho_named_'))/(safe_name+'.pdf')
        named_excel = named_pdf.with_suffix('.xlsx')
        named_pdf.write_bytes(project_pdf_bytes)
        named_excel.write_bytes(project_excel_bytes)
        return str(named_pdf),str(named_excel)
    if survey_pdf is None or response_xlsx is None:
        raise ValueError('Please upload both survey PDF and Excel.')
    return save_upload(survey_pdf), save_upload(response_xlsx)

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
        if not attend_input:
            st.error('Please upload attendance.')
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
                pdf_path, xlsx_path = selected_survey_files()
                with st.spinner('Generating your evaluation report...'):
                    engine.build(report_template,pdf_path,xlsx_path,str(report),attendance_people=people,attendance_days=dcount)
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

st.caption('Private survey exports are retrieved from your configured private GitHub Project Library. Attendance files are processed by the application host. Keep the website access restricted.')
