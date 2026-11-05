# Google Drive templates: store these two files ONCE in MyDrive/Evaluation_Templates
import pathlib

DRIVE_TEMPLATE_FOLDER = "Evaluation_Templates"  # relative to MyDrive
REPORT_TEMPLATE_FILENAME = "report_reference.pptx"
ATTENDANCE_TEMPLATE_FILENAME = "attendance_reference.pptx"

def install_embedded_templates():
    """Preserved call signature: load templates from Google Drive without embedding them."""
    from google.colab import drive
    drive.mount('/content/drive')
    folder = pathlib.Path('/content/drive/MyDrive') / DRIVE_TEMPLATE_FOLDER
    report = folder / REPORT_TEMPLATE_FILENAME
    attendance = folder / ATTENDANCE_TEMPLATE_FILENAME
    missing = [str(p) for p in (report, attendance) if not p.is_file()]
    if missing:
        raise FileNotFoundError(
            "Save both template PPTX files in Google Drive under MyDrive/"
            + DRIVE_TEMPLATE_FOLDER + ". Missing: " + ", ".join(missing)
        )
    print('Original PPTX templates loaded from Google Drive; no template uploads needed.')
    return str(report), str(attendance)

# Auto Evaluation + Attendance | template-preserving report, striped attendance
# Charts retain exact template wording and styling. Values show whole percentages.
"""Evaluation report with fixed templates stored in Google Drive (no template uploads)."""
import re, json, zipfile, pathlib, collections, statistics, copy
from datetime import datetime
import pymupdf
from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE

SCORE={'weak':20,'poor':20,'average':40,'fair':40,'good':60,'very good':80,'excellent':100}
SECTIONS=[('Learning Outcomes','Learning Outcome:',4),('Workshop Content','The Workshop Content:',5),('Trainer','Instructor:',4)]
COMMENT_KEYS=['learning outcome\'s ratings','workshop content\'s ratings','instructor\'s ratings','most interesting','changes do you suggest','other simulations']
COMMENT_TITLES=['Participants’ comments on the learning outcomes:', 'Participants’ comments on the workshop content:', 'Participants’ comments on the trainer:', 'What participants found most interesting:', 'Suggested changes for improvement:', 'Other simulations / topics requested:']

def safe_text(sh, value):
    """Replace characters in existing runs without deleting run formatting or paragraph layout."""
    if not sh.has_text_frame: return
    paragraphs=sh.text_frame.paragraphs
    if not paragraphs:return
    runs=[r for p in paragraphs for r in p.runs]
    if not runs: return
    runs[0].text=str(value)
    for r in runs[1:]: r.text=''

def recurse(shapes):
    for s in shapes:
        yield s
        if s.shape_type==MSO_SHAPE_TYPE.GROUP:
            yield from recurse(s.shapes)

def remove_shape(sh):
    el=sh._element;el.getparent().remove(el)

def pdf_data(pdf_path):
    doc=pymupdf.open(pdf_path)
    pages=[page.get_text(sort=True) for page in doc]
    full='\n'.join(pages)
    first=pages[0].splitlines()
    heading=next((t.strip() for t in first if re.search(r' - .* - ', t) and not t.startswith('http')), '')
    title=heading if heading and len(heading)<220 else pathlib.Path(pdf_path).stem
    categories={}
    for label,pat in [('Learning Outcomes',r'Learning Outcome:'),('Workshop Content',r'The Workshop Content:'),('Trainer',r'Instructor:')]:
        match=re.search(pat+r'\s*(\d{2,3}\.\d{1,2})\s+100',full,re.I)
        if match:categories[label]=float(match.group(1))
    overall=re.search(r'Overall average survey rating\s*:\s*(\d+(?:\.\d+)?)\s*/\s*10',full,re.I)
    summary={}
    for lab in ['Survey Visits','Total Responses','Completed Responses','Partial Responses']:
        m=re.search(r'(\d+)\s*\n\s*'+re.escape(lab),full,re.I)
        if m:summary[lab]=int(m.group(1))
    # Individual question ratings: extract the last percentage adjacent to each 100-point survey prompt.
    section_scores={}
    for label,code,n in SECTIONS:
        # Each section's final "Average rating" terminates the PDF question block.
        if label=='Learning Outcomes':pattern=r'Q4\s+Learning Outcome:'
        elif label=='Workshop Content':pattern=r'Q5\s+The Workshop Content:'
        else:pattern=r'Q6\s+Instructor:'
        start=re.search(pattern,full,re.I)
        if not start:continue
        match=re.search(r'Average rating\s*:\s*\d+(?:\.\d+)?',full[start.end():],re.I)
        block=full[start.end():start.end()+match.end()] if match else full[start.end():]
        # Some PDFs print tabular average right before response count; prefer those scores when usable.
        found=re.findall(r'(?<!\d)(\d{2,3}\.\d{2})\s+\d{1,2}(?:\s|$)',block)
        vals=[float(v) for v in found if 50<=float(v)<=100]
        if len(vals)>=n:section_scores[label]=vals[:n]
    doc.close()
    return {'title':title,'category_scores':categories,'overall':float(overall.group(1))*10 if overall else None,
            'survey_counts':summary,'question_scores':section_scores}

def meaningful_comment(value):
    """Exclude empty, placeholder, punctuation-only and non-substantive feedback."""
    if value is None:
        return False
    original = str(value).strip()
    normalized = re.sub(r"[\s\W_]+", " ", original.casefold(), flags=re.UNICODE).strip()
    reject = {
        '', 'none', 'null', 'nil', 'na', 'n a', 'n a n a', 'not applicable',
        'no response', 'no comment', 'no comments', 'no feedback',
        'no suggestion', 'no suggestions', 'nothing', 'nothing to add',
        'nothing else', 'nothing for now', 'no changes', 'no change',
        'no improvement', 'no improvements', 'no idea', 'not sure',
        'dont know', 'don t know', 'idk', 'all good', 'all fine',
        'everything is fine', 'everything is good', 'everything was good',
        'okay', 'ok', 'yes', 'no', 'thanks', 'thank you', 'test',
        'لا يوجد', 'لا شيء', 'مفيش', 'لا', 'لا تعليق', 'لا توجد', 'لا يوجد تعليق',
    }
    if normalized in reject:
        return False
    if len(re.sub(r"[^\w]+", "", normalized, flags=re.UNICODE)) < 3:
        return False
    # One-letter filler, e.g., xxxxx or "---".
    if len(set(normalized.replace(' ', ''))) == 1:
        return False
    return True


def arrange_comment_boxes(slide, groups, items):
    """Pair each original card BACKGROUND with its text group, then move/delete BOTH.

    The template stores six card backgrounds as independent auto-shapes, not as
    children of the text groups. Moving a group alone leaves text outside its box.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    count = min(len(items), 6)
    groups = list(groups)
    if len(groups) != 6:
        raise ValueError(f'Expected six comment text groups; found {len(groups)}')

    # Card backgrounds are large blank rectangles immediately behind the six
    # text groups. Restrict to the known six-card grid so other decorations stay.
    backgrounds = [sh for sh in slide.shapes
                   if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                   and 2.7 <= sh.width / 914400 <= 4.2
                   and 1.8 <= sh.height / 914400 <= 2.8
                   and 1.5 < sh.top / 914400 < 5.5]
    if len(backgrounds) != 6:
        raise ValueError(f'Expected six separate comment-card backgrounds; found {len(backgrounds)}')

    backgrounds.sort(key=lambda sh: (round(sh.top / 914400, 1), sh.left))
    # Original 3x2 grid in reading order; use matching center proximity to pair.
    available = groups[:]
    cards = []
    for bg in backgrounds:
        cx, cy = bg.left + bg.width//2, bg.top + bg.height//2
        grp = min(available, key=lambda g: abs((g.left + g.width//2)-cx) + abs((g.top + g.height//2)-cy))
        available.remove(grp)
        cards.append((bg, grp))

    top, bottom = cards[:3], cards[3:]
    xs = [int(bg.left) for bg,_ in top]
    ys = [int(top[0][0].top), int(bottom[0][0].top)]
    w = int(top[0][0].width)
    h = int(top[0][0].height)
    center_x = (xs[0] + xs[2] + w)//2
    center_y = (ys[0] + ys[1] + h)//2
    def row_xs(n):
        if n == 1: return [center_x - w//2]
        if n == 2: return [(xs[0]+xs[1])//2, (xs[1]+xs[2])//2]
        return xs
    coordinates = {
        0: [],
        1: [(row_xs(1)[0],center_y-h//2)],
        2: [(x,center_y-h//2) for x in row_xs(2)],
        3: [(x,center_y-h//2) for x in row_xs(3)],
        4: [(x,y) for y in ys for x in row_xs(2)],
        5: [(x,ys[0]) for x in row_xs(3)] + [(x,ys[1]) for x in row_xs(2)],
        6: [(x,y) for y in ys for x in row_xs(3)],
    }
    for i,(bg,grp) in enumerate(cards):
        if i >= count:
            remove_shape(grp)
            remove_shape(bg)
            continue
        target_x,target_y = coordinates[count][i]
        dx,dy = int(target_x-bg.left),int(target_y-bg.top)
        # Move background and foreground by exactly the same displacement.
        bg.left = int(bg.left + dx)
        bg.top = int(bg.top + dy)
        xfrm = grp._element.grpSpPr.xfrm
        xfrm.off.x = int(xfrm.off.x + dx)
        xfrm.off.y = int(xfrm.off.y + dy)
        leaves = [sh for sh in recurse(grp.shapes)
                  if sh.has_text_frame and sh.text.strip() in ('Comment','Name')]
        for sh in leaves:
            is_comment = sh.text.strip() == 'Comment'
            value = items[i][0] if is_comment else items[i][1]
            if is_comment: value = '“' + value.strip().strip('"“”') + '”'
            safe_text(sh,value)

def xlsx_data(path):
    wb=load_workbook(path,read_only=True,data_only=True)
    ws=wb['Responses'] if 'Responses' in wb.sheetnames else wb.active
    rows=ws.iter_rows(values_only=True)
    head=next((r for r in rows if any(str(x or '').strip().lower()=='response id' for x in r)),None)
    if head is None: raise ValueError('Cannot locate the Responses sheet header (Response ID).')
    def norm(v):return re.sub(r'\s+',' ',str(v or '')).strip().lower()
    headers=[norm(v) for v in head]
    namecol=next((i for i,h in enumerate(headers) if h in ('name','name:','participant name')),None)
    ratings=collections.defaultdict(list);comments=collections.defaultdict(list)
    count=0
    for row in rows:
        if not any(v is not None and str(v).strip() for v in row):continue
        count+=1
        person=str(row[namecol]).strip() if namecol is not None and row[namecol] is not None else 'Anonymous participant'
        for i,h in enumerate(headers):
            val=row[i] if i<len(row) else None
            if val is None or str(val).strip()=='':continue
            v=str(val).strip()
            group=next((g for g,keyword,_ in SECTIONS if h.startswith(keyword.lower())),None)
            if group and '[comments' not in h:
                if v.lower() in SCORE: ratings[group].append((h, SCORE[v.lower()]))
            for j,k in enumerate(COMMENT_KEYS):
                if k in h and meaningful_comment(v):
                    if (v.casefold(),person.casefold()) not in [(a.casefold(),b.casefold()) for a,b in comments[j]]:
                        comments[j].append((v,person))
                    break
    # Generate individual question means by header, in original XLSX column order.
    questions={}
    for group,_,_ in SECTIONS:
        cols=[]
        for h in headers:
            if h.startswith(next(keyword.lower() for g,keyword,_ in SECTIONS if g==group)) and '[comments' not in h:
                vals=[v for header,v in ratings[group] if header==h]
                if vals:cols.append((h,round(statistics.mean(vals),2)))
        questions[group]=cols
    return {'respondents_in_excel':count,'questions':questions,'comments':dict(comments)}

def set_chart(slide, category, vals, question_labels=None):
    """Update numeric values only; keep template's exact category labels and chart styling."""
    sh = next((s for s in slide.shapes if s.has_chart), None)
    if sh is None or not vals:
        return False
    chart = sh.chart
    original_labels = [c.label for c in chart.plots[0].categories]
    if len(vals) != len(original_labels):
        raise ValueError(f"{category}: expected {len(original_labels)} values, received {len(vals)}")
    original_series = chart.series[0].name
    cd = CategoryChartData()
    cd.categories = original_labels  # DO NOT replace text with Excel survey headers.
    cd.add_series(original_series, [round(float(v) / 100, 6) for v in vals])
    chart.replace_data(cd)
    # Replace_data may keep old point-label number formats. Force any existing
    # per-point dLbl numFmt to whole percentages without changing label layout.
    for node in chart._chartSpace.xpath('.//c:dLbls/c:dLbl/c:numFmt'):
        node.set('formatCode', '0%')
        node.set('sourceLinked', '0')
    # Template has percentages stored as fractions, e.g., .94 = 94%.
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = '0%'
    plot.data_labels.number_format_is_linked = False
    # Important: the number printed DIRECTLY ABOVE EACH BAR (not only the Y-axis).
    # Keep the original position, font, color, and chart-category wording.
    # All charts store .97 as fractions, so 0% displays 97%.
    from pptx.oxml.xmlchemy import OxmlElement
    for labels_xml in chart._chartSpace.xpath('.//c:dLbls'):
        fmts = labels_xml.xpath('./c:numFmt')
        if fmts:
            fmt = fmts[0]
        else:
            fmt = OxmlElement('c:numFmt')
            # c:numFmt must come after c:dLbl entries, before c:spPr/c:txPr
            ref = next((node for node in labels_xml
                        if node.tag.rsplit('}', 1)[-1] in
                        ('spPr','txPr','dLblPos','showLegendKey','showVal')),
                       None)
            if ref is not None:
                labels_xml.insert(list(labels_xml).index(ref), fmt)
            else:
                labels_xml.append(fmt)
        fmt.set('formatCode', '0%')
        fmt.set('sourceLinked', '0')
        for point_fmt in labels_xml.xpath('./c:dLbl/c:numFmt'):
            point_fmt.set('formatCode','0%')
            point_fmt.set('sourceLinked','0')
    if chart.has_legend:
        pass  # Do not change legend appearance.
    try:
        chart.value_axis.tick_labels.number_format = '0%'
        chart.value_axis.tick_labels.number_format_is_linked = False
    except (AttributeError, ValueError):
        pass
    return True


def ordinal_day(day):
    suffix = 'th' if 11 <= day % 100 <= 13 else {1:'st',2:'nd',3:'rd'}.get(day % 10,'th')
    return f'{day}{suffix}'


def pretty_training_date(raw):
    """Accept day-month-year or month-day-year when one component exceeds 12."""
    if not raw:
        return 'Not provided'
    text = str(raw).strip()
    # Day span like 4-5-10-2026 (4th–5th October 2026)
    m = re.fullmatch(r'(\d{1,2})[-/](\d{1,2})[-/](\d{1,2})[-/](\d{4})',text)
    if m:
        d1,d2,month,year = map(int,m.groups())
        if all(1<=d<=31 for d in [d1,d2]) and 1<=month<=12:
            return f'{ordinal_day(d1)} – {ordinal_day(d2)} {datetime(year,month,1).strftime("%B %Y")}'
    m = re.fullmatch(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})',text)
    if m:
        a,b,year = map(int,m.groups())
        if a>12 and b<=12: day,month=a,b
        elif b>12 and a<=12: day,month=b,a
        else: day,month=a,b # default European order
        try:
            dt=datetime(year,month,day)
            return f'{ordinal_day(day)} {dt.strftime("%B %Y")}'
        except ValueError:
            return text
    return text


def replace_labeled_paragraph(paragraph, label, value):
    """Retain the template's label run and bold value runs, without resetting text frames."""
    runs=list(paragraph.runs)
    if not runs: return False
    # Label stays in original run, with original non-bold formatting.
    runs[0].text=label + ' '
    if len(runs)==1:
        value_run=paragraph.add_run()
        value_run.font.name=runs[0].font.name
        value_run.font.size=runs[0].font.size
    else:
        value_run=runs[1]
    value_run.text=str(value)
    value_run.font.bold=True
    for run in runs[2:]: run.text=''
    return True


def update_workshop_detail(shape, updates):
    if not shape.has_text_frame: return False
    changed=False
    for p in shape.text_frame.paragraphs:
        for label, value in updates.items():
            if p.text.strip().startswith(label):
                replace_labeled_paragraph(p,label,value)
                changed=True
                break
    return changed

def remove_slides(prs, indices):
    for idx in sorted(indices,reverse=True):
        sid=prs.slides._sldIdLst[idx]
        prs.part.drop_rel(sid.rId)
        prs.slides._sldIdLst.remove(sid)

def derive_metadata(title):
    """Filename: Client - Training Name[_Run_Day] - Trainer - D[-D]-M-YYYY.

    Parse from the RIGHT so training names can contain extra hyphens.
    """
    raw = pathlib.Path(str(title)).stem
    raw = re.sub(r'\s*\(\d+\)$', '', raw).strip()  # Colab duplicated filename suffix
    # Supports 4-5-10-2026 / 02-03-10-2026 / 14-9-2026
    date_match = re.search(r'(?:^|\s*[-–]\s*)(\d{1,2}[-/]\d{1,2}(?:[-/]\d{1,2})?[-/]\d{4})\s*$', raw)
    date_value = date_match.group(1) if date_match else None
    leading = raw[:date_match.start()].strip(' -–_') if date_match else raw
    components = [v.strip(' _') for v in re.split(r'\s+[-–]\s+',leading) if v.strip(' _')]
    client = components[0] if components else 'Unspecified client'
    trainer = components[-1] if len(components) >= 3 else 'Not provided'
    training = components[1:-1] if len(components)>=3 else components[1:]
    # Some exports concatenate the last training segment and trainer without spaces:
    # '...-Group 6-Virtual-Rasha Abu Elenen-09-15-2026'.
    if trainer == 'Not provided' and len(components)==2 and '-' in components[1]:
        prefix, candidate = components[1].rsplit('-',1)
        if re.fullmatch(r'[A-Za-zÀ-ÿ]+(?:\s+[A-Za-zÀ-ÿ]+){1,5}',candidate.strip()):
            trainer = candidate.strip()
            training = [prefix.strip()]
    course = ' - '.join(training).strip()
    # Keep Run N as part of workshop title; normalize underscore separators.
    course = re.sub(r'\s*_+\s*', ' - ', course).strip()
    return {'client':client,'trainer':trainer,'course':course or 'Not provided',
            'date':date_value,'run':next((x for x in training if re.search(r'Run\s*\d+',x,re.I)), '')}


def update_trainer_chart_subtitle(slide, trainer_name):
    """Replace only the template trainer subtitle, keeping its original text runs."""
    if not trainer_name or trainer_name == 'Not provided':
        return False
    for shape in recurse(slide.shapes):
        if not shape.has_text_frame:
            continue
        paras = shape.text_frame.paragraphs
        # Template uses a two-paragraph title: 'Trainer' and 'Eng: Gehad Elsobky'.
        if not any(p.text.strip().casefold() == 'trainer' for p in paras):
            continue
        for idx, p in enumerate(paras):
            if p.text.strip().casefold() == 'trainer' and idx + 1 < len(paras):
                subtitle = paras[idx + 1]
                if not subtitle.text.strip():
                    continue
                if subtitle.runs:
                    subtitle.runs[0].text = trainer_name
                    for run in subtitle.runs[1:]:
                        run.text = ''
                else:
                    subtitle.add_run().text = trainer_name
                return True
    return False


def workshop_day_count(raw, attendance_days):
    # Dates like 2-3-10-2026 explicitly mean two days.
    if raw and re.fullmatch(r'\d{1,2}[-/]\d{1,2}[-/]\d{1,2}[-/]\d{4}',str(raw)):
        a,b,month,year=map(int,re.split(r'[-/]',raw))
        if 1<=a<=31 and 1<=b<=31 and 1<=month<=12:
            return abs(b-a)+1
    # Date metadata often indicates one day; attendance has direct day-column evidence.
    return max(1, int(attendance_days))


def update_workshop_metrics(detail, days, planned, actual):
    """Replace four metric VALUES by their visual left-to-right order, not by stale numbers."""
    metric_values=[]
    for sh in detail.shapes:
        if sh.has_text_frame and re.fullmatch(r'\s*(?:\d+|N/A)\s*', sh.text):
            metric_values.append(sh)
    metric_values.sort(key=lambda sh:sh.left)
    if len(metric_values)!=4:
        raise ValueError(f'Expected four workshop metric values, found {len(metric_values)}; template layout differs.')
    for sh, val in zip(metric_values,[days,days*8,planned,actual]):
        safe_text(sh,str(val))


def build(template,pdf,xlsx,out,attendance_people=None, attendance_days=None):
    info=pdf_data(pdf);excel=xlsx_data(xlsx);meta=derive_metadata(pathlib.Path(pdf).stem)
    prs=Presentation(template);audit=[]
    def log(msg):audit.append(msg);print(msg)
    # Keep the original template slide masters, backgrounds, fonts and named shapes.
    for slide_idx in [0,1]:
        for s in recurse(prs.slides[slide_idx].shapes):
            if s.has_text_frame and 'Evaluation' in s.text and 'Report' in s.text:
                # Preserve heading's existing Barlow formatting.
                safe_text(s,'Evaluation\nReport')
    detail=prs.slides[3]
    for sh in recurse(detail.shapes):
        update_workshop_detail(sh, {
            'Trainer’s Name:': meta['trainer'],
            'Workshop Title:': meta['course'],
            'Date:': pretty_training_date(meta['date']),
            'Location:': 'Not provided',
        })
    if attendance_people is None or attendance_days is None:
        raise ValueError('Attendance data must be parsed before generating workshop metrics.')
    days = workshop_day_count(meta['date'], attendance_days)
    planned = len(attendance_people)
    # Distinct people who attended at least one workshop day.
    actual = sum('present' in statuses for _, statuses in attendance_people)
    update_workshop_metrics(detail, days, planned, actual)
    log(f'Workshop metrics: {days} day(s), {days*8} hours, {planned} planned, {actual} actually attended (at least one day).')
    for idx,grp,n in [(5,'Learning Outcomes',4),(6,'Workshop Content',5),(7,'Trainer',4)]:
        # XLSX response-by-response data is the authoritative source for individual bars;
        # PDF text extraction frequently changes visual bar order.
        vals=[v for _,v in excel['questions'].get(grp,[])]
        log('Calculated individual question chart from Excel responses for '+grp+'.')
        labels=[q for q,v in excel['questions'].get(grp,[])]
        if len(vals)==n: set_chart(prs.slides[idx],grp,vals,labels)
        else: log('WARNING insufficient valid scores for '+grp+'; chart must be reviewed.')
    if not update_trainer_chart_subtitle(prs.slides[7], meta['trainer']):
        log('WARNING trainer subtitle on chart slide could not be updated from template.')
    # Update grouped Feedback numbers while preserving the number boxes and fonts.
    feedback=prs.slides[8]
    targets=[info['category_scores'].get('Learning Outcomes'),info['category_scores'].get('Workshop Content'),
             info['category_scores'].get('Trainer'),info['overall']]
    values=[s for s in recurse(feedback.shapes) if s.has_text_frame and re.fullmatch(r'\s*\d+(?:\.\d+)?%\s*',s.text)]
    for sh,num in zip(values,targets):
        if num is not None:safe_text(sh,f'{num:.1f}%')
    # Populate six comment layouts; preserve frame, group, logos, headers and colors.
    for section in range(6):
        slide=prs.slides[12+section]
        subtitle=next((s for s in slide.shapes if s.has_text_frame and 'comments on the learning outcome' in s.text.lower()),None)
        if subtitle:safe_text(subtitle,COMMENT_TITLES[section])
        items=excel['comments'].get(section,[])
        # Keep 6 substantive, nonduplicated comments, longest first, stable tie break.
        items=sorted(items,key=lambda v:-len(v[0]))[:6]
        # Only the six editable comment-card groups, never decorative/logo groups.
        groups=[]
        for sh in slide.shapes:
            if sh.shape_type != MSO_SHAPE_TYPE.GROUP:
                continue
            leaves=[x.text.strip() for x in recurse(sh.shapes) if x.has_text_frame]
            if 'Comment' in leaves and 'Name' in leaves:
                groups.append(sh)
        arrange_comment_boxes(slide, groups, items)
        log(f'{COMMENT_TITLES[section]} {len(items)} of {len(excel["comments"].get(section,[]))} comments selected.')
    # Remove guidelines-only instruction pages, unsupported pre/post slides and empty media pages.
    # Keep the original Appendix at index 2; remove only 0,9,10,18,19.
    remove_slides(prs,[0,9,10,18,19])
    prs.save(out)
    report={'source_pdf':str(pdf),'source_excel':str(xlsx),'source_template':str(template),
            'metadata':meta,'extracted_pdf':info,'excel_response_count':excel['respondents_in_excel'],
            'warnings':audit,'output':str(out)}
    report_path=str(pathlib.Path(out).with_suffix('.audit.json'))
    pathlib.Path(report_path).write_text(json.dumps(report,indent=2,ensure_ascii=False),encoding='utf-8')
    return out, report_path



# ---------- One-slide attendance, retaining original template style ----------
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.image import ImagePart
from pptx.oxml.ns import qn
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from openpyxl.styles.colors import COLOR_INDEX
from PIL import Image, ImageDraw
from datetime import date
from io import BytesIO
from copy import deepcopy
import math

REPORT_PPTX = 'Evaluation_Report_Automated.pptx'
OUTPUT_PPTX = 'Evaluation_Report_With_Attendance_FINAL.pptx'
MAX_DAY_COLUMNS = 2
NAME_COLUMN = None
DAY_COLUMNS = None
BLANK_AS_ABSENT = False
ZEBRA_COLOR = RGBColor(67, 82, 100)

def cell_text(x): return str(x).strip() if x is not None else ''

def color_rgb(cell):
    color=cell.fill.fgColor
    if cell.fill.patternType!='solid': return None
    if color.type=='rgb' and color.rgb:
        try: return tuple(bytes.fromhex(str(color.rgb)[-6:]))
        except: return None
    if color.type=='indexed' and color.indexed is not None:
        try: return tuple(bytes.fromhex(COLOR_INDEX[color.indexed][-6:]))
        except: return None
    return None

def mark(cell):
    val=cell_text(cell.value).casefold()
    if val in {'attended','present','yes','y','true','1','p','✅','✔','✓','☑','حاضر'}:return 'present'
    if val in {'absent','no','n','false','0','a','❌','✗','✘','غائب'}:return 'absent'
    rgb=color_rgb(cell)
    if rgb:
        r,g,b=rgb
        if g>=85 and g>r*1.20 and g>b*1.15:return 'present'
        if r>=100 and r>g*1.25 and r>b*1.25:return 'absent'
    if val in {'','-','—','na','n/a','none'}:return 'absent' if BLANK_AS_ABSENT else 'unknown'
    return 'unknown'

def parse_sheet(ws):
    header=None
    for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,25)):
        for c in row:
            if cell_text(c.value).casefold() in {'name','participant name','full name','employee name','attendee','اسم'}:
                header=(c.row,c.column);break
        if header:break
    if header:
        hr,nc=header
    else:
        # Headerless color-coded attendance sheets: names usually start in column A, row 2.
        nc=NAME_COLUMN or 1
        hr=1
        # If first row is already a name and has colored attendance cells, treat row 1 as data.
        if cell_text(ws.cell(1,nc).value) and any(mark(ws.cell(1,c))!='unknown' for c in range(nc+1,min(ws.max_column,nc+3)+1)):
            hr=0
    if NAME_COLUMN: nc=NAME_COLUMN
    if DAY_COLUMNS is not None: days=list(DAY_COLUMNS)
    else:
        days=[]
        for col in range(1,ws.max_column+1):
            if col==nc:continue
            h=ws.cell(max(1,hr),col).value
            s=cell_text(h).casefold()
            is_date=isinstance(h,(datetime,date)) or bool(re.search(r'\bday\s*\d+\b|\b\d{1,2}[-/]\d{1,2}\b|\b\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b',s))
            examples=[mark(ws.cell(r,col)) for r in range(max(1,hr+1),min(ws.max_row,hr+30)+1)]
            hits=sum(x!='unknown' for x in examples)
            if is_date or hits>=2:days.append(col)
        if not days and ws.max_column>nc:
            # Color-coded spreadsheets without an obvious header: consider adjacent cols, only if color-coded.
            days=[c for c in range(nc+1,ws.max_column+1) if any(color_rgb(ws.cell(r,c)) for r in range(2,min(ws.max_row,20)+1))]
    if not days:return None
    if len(days)>MAX_DAY_COLUMNS:
        raise ValueError(f'{ws.title}: Found {len(days)} attendance columns. Reference has only 2; choose DAY_COLUMNS manually or use a wider table.')
    names=[]
    for r in range(hr+1,ws.max_row+1):
        name=cell_text(ws.cell(r,nc).value)
        if not name or name.casefold() in {'total','name','participant name','total attendees'}:continue
        statuses=[mark(ws.cell(r,c)) for c in days]
        if all(v=='unknown' for v in statuses) and (not re.search('[a-zA-Z\u0600-\u06ff]',name) or len(name)>120):continue
        names.append((name,statuses))
    if not names:return None
    labels=[]
    for i,c in enumerate(days):
        h=ws.cell(max(hr,1),c).value if hr else None
        labels.append(h.strftime('%d %b') if isinstance(h,(date,datetime)) else cell_text(h) or f'Day {i+1}')
    return names,labels,ws.title


def change_attendance_text(cell, value):
    frame=cell.text_frame
    p=frame.paragraphs[0]
    if not p.runs:p.add_run()
    p.runs[0].text=str(value)
    for r in p.runs[1:]:r.text=''
    for q in frame.paragraphs[1:]:
        for r in q.runs:r.text=''

def attendance_symbol(kind):
    s=160; img=Image.new('RGBA',(s,s),(0,0,0,0));d=ImageDraw.Draw(img)
    if kind=='present':
        d.rounded_rectangle((12,12,148,148),radius=26,fill=(24,169,78,255))
        d.line([(36,81),(66,109),(122,46)],fill='white',width=15,joint='curve')
    elif kind=='absent':
        d.rounded_rectangle((12,12,148,148),radius=26,fill=(217,55,65,255))
        d.line([(47,47),(113,113)],fill='white',width=14)
        d.line([(113,47),(47,113)],fill='white',width=14)
    else:
        d.line([(35,80),(125,80)],fill=(135,140,150,255),width=12)
    f=BytesIO();img.save(f,format='PNG');return f.getvalue()

def make_attendance_one_slide(attendance_xlsx, attendance_template, output='Attendance_Only_One_Slide_Striped.pptx', parsed=None):
    if parsed is not None:
        people,day_labels,sheet_name=parsed
    else:
        wb=load_workbook(attendance_xlsx,data_only=True)
        candidates=[]
        for ws in wb.worksheets:
            found=parse_sheet(ws)
            if found:candidates.append(found)
        if not candidates:raise ValueError('Cannot find attendance records in the Excel file.')
        people,day_labels,sheet_name=max(candidates,key=lambda x:len(x[0]))
    print('Attendance:', sheet_name, '|',len(people),'participants |', day_labels)
    prs=Presentation(attendance_template)
    index=next((i for i,s in enumerate(prs.slides) if len([x for x in s.shapes if x.has_table and len(x.table.columns)==3])>=2),None)
    if index is None:raise ValueError('Attendance reference must have two 3-column editable tables.')
    # Leave exactly one original attendance slide: no rebuilding, original theme/media intact.
    for i in reversed(range(len(prs.slides))):
        if i!=index:
            sid=prs.slides._sldIdLst[i];prs.part.drop_rel(sid.rId);prs.slides._sldIdLst.remove(sid)
    slide=prs.slides[0]
    tables=sorted([s for s in slide.shapes if s.has_table and len(s.table.columns)==3],key=lambda x:x.left)[:2]
    count_left=math.ceil(len(people)/2);counts=[count_left,len(people)-count_left]
    if max(counts)>38:raise ValueError('Over 76 participants: cannot fit legibly on one slide.')
    bottom=int(prs.slide_height-Inches(0.35))
    marks=[];cursor=0
    for sh,needed in zip(tables,counts):
        tab=sh.table
        if len(tab.rows) == 0:
            raise ValueError('The attendance reference contains a table with no rows. Use the original formatted attendance PPTX.')
        while len(tab.rows)<needed+1:tab._tbl.append(deepcopy(tab.rows[len(tab.rows)-1]._tr))
        while len(tab.rows)>needed+1:tab._tbl.remove(tab.rows[len(tab.rows)-1]._tr)
        total=len(tab.rows)
        height=min(int(Inches(.28)),(bottom-int(sh.top))//total)
        if height<int(Inches(.145)):raise ValueError('Too many names to fit visibly on one slide.')
        for row in tab.rows:row.height=height
        fontsize=max(8,min(13,(height/914400)*43))
        change_attendance_text(tab.cell(0,0),'Name')
        for d in range(2):change_attendance_text(tab.cell(0,d+1),day_labels[d] if d<len(day_labels) else '')
        for ri in range(1,total):
            name,statuses=people[cursor];cursor+=1
            cell=tab.cell(ri,0);change_attendance_text(cell,name)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            cell.margin_top=0;cell.margin_bottom=0;cell.margin_left=Inches(.045);cell.margin_right=Inches(.025)
            for p in cell.text_frame.paragraphs:
                p.alignment=PP_ALIGN.LEFT
                for run in p.runs:run.font.size=Pt(fontsize if len(name)<=33 else max(7,fontsize*(33/len(name))**.43))
            # Alternating stripes across the WHOLE row, in both tables.
            for ci in range(3):
                c=tab.cell(ri,ci)
                if ri%2:
                    c.fill.solid();c.fill.fore_color.rgb=ZEBRA_COLOR
                else:c.fill.background()
            y=int(sh.top)+ri*height;x=int(sh.left)+int(tab.columns[0].width)
            for d in range(2):
                c=tab.cell(ri,d+1);change_attendance_text(c,'')
                if d<len(statuses):
                    width=int(tab.columns[d+1].width)
                    sz=int(min(height*.68,width*.30))
                    marks.append((statuses[d],x+(width-sz)//2,y+(height-sz)//2,sz))
                x+=int(tab.columns[d+1].width)
    icons={k:attendance_symbol(k) for k in ['present','absent','unknown']}
    for kind,x,y,sz in marks:slide.shapes.add_picture(BytesIO(icons[kind]),x,y,width=sz,height=sz)
    prs.save(output)
    verify=Presentation(output);written=[]
    for sh in verify.slides[0].shapes:
        if sh.has_table and len(sh.table.columns)==3:
            written.extend(sh.table.cell(i,0).text.strip() for i in range(1,len(sh.table.rows)))
    if len(verify.slides)!=1 or written!=[name for name,_ in people]:
        raise RuntimeError('Attendance names did not match the source workbook after saving.')
    return output, len(people)


# ---------- Scanned paper attendance (image mode) ----------
# Scan layouts: optional # column, name column, 1-2 dated signature columns.
# Blue-ink signatures are separated from printed black grid lines.
# Image parsing is reviewed before any report is generated.

def _unique_positions(values, tolerance=12):
    groups=[]
    for x in sorted(map(int,values)):
        if not groups or x-groups[-1][-1]>tolerance: groups.append([x])
        else: groups[-1].append(x)
    return [round(sum(group)/len(group)) for group in groups]

def _extract_grid(image):
    import cv2, numpy as np
    from PIL import ImageOps
    from PIL import Image as PILImage
    original=PILImage.open(image)
    original=ImageOps.exif_transpose(original).convert('RGB')
    arr=cv2.cvtColor(np.asarray(original),cv2.COLOR_RGB2BGR)
    h,w=arr.shape[:2]
    # Use deskewed crop boundaries: avoid requiring exact 90-degree photographs.
    gray=cv2.cvtColor(arr,cv2.COLOR_BGR2GRAY)
    adaptive=cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV,35,13)
    hor=cv2.morphologyEx(adaptive,cv2.MORPH_OPEN,
                         cv2.getStructuringElement(cv2.MORPH_RECT,(max(45,w//4),1)))
    ver=cv2.morphologyEx(adaptive,cv2.MORPH_OPEN,
                         cv2.getStructuringElement(cv2.MORPH_RECT,(1,max(45,h//7))))
    # Count pixels per scanline so weak partial photo lines do not disappear.
    ys=_unique_positions(np.where(np.count_nonzero(hor,axis=1)>w*.20)[0],max(7,h//160))
    xs=_unique_positions(np.where(np.count_nonzero(ver,axis=0)>h*.26)[0],max(7,w//120))
    if len(xs)<4 or len(ys)<4:
        raise ValueError('Cannot detect the attendance grid. Retake the photo straight-on with all table borders visible, or use Excel mode.')
    # Number, Name, day1, day2: 5 boundaries. 1-day image: 4.
    # Keep the large table border lines, not the lines of handwriting.
    return arr,gray,sorted(xs),sorted(ys)

def read_attendance_image(image_path, name_reference=None):
    """Return (people, day_labels, sheet_name), with manual review recommended.

    A mark is evidence of visible ink, not evidence of who actually signed.
    Images must be reviewed before committing the output.
    """
    import cv2, numpy as np, pytesseract
    arr,gray,xs,ys=_extract_grid(image_path)
    # Usually grid edges x0=#, x1=Name, x2=Day1, x3=Day2, x4=end.
    # Handle unnumbered tables having Name, Day1, Day2 (4 edges).
    if len(xs)>5:
        # select five best column edges via largest continuous ver lines
        # do not silently choose if table structure is unusual
        raise ValueError(f'Found {len(xs)} vertical dividers instead of 4 or 5. Please use a straight, cropped image of the attendance table.')
    if len(xs)==5:
        name_bounds=(xs[1],xs[2]); day_bounds=[(xs[2],xs[3]),(xs[3],xs[4])]
    elif len(xs)==4:
        name_bounds=(xs[0],xs[1]);day_bounds=[(xs[1],xs[2]),(xs[2],xs[3])]
    else:
        raise ValueError('Cannot identify name and date columns in this image.')
    if len(ys)>100: raise ValueError('Too many horizontal lines detected; crop the image to the attendance table.')
    header_top,header_bottom=ys[0],ys[1]
    labels=[]
    for d,(x0,x1) in enumerate(day_bounds):
        crop=gray[header_top+3:header_bottom-3,x0+3:x1-3]
        raw=pytesseract.image_to_string(crop,config='--psm 7').strip()
        labels.append(re.sub(r'\s+',' ',raw) or f'Day {d+1}')
    hsv=cv2.cvtColor(arr,cv2.COLOR_BGR2HSV)
    # Signature detection for blue/purple ballpoint on white paper. Also detect
    # dark non-grid handwriting, but mark weak findings UNKNOWN for review.
    candidates=[]
    for idx in range(1,len(ys)-1):
        y0,y1=ys[idx:idx+2]
        if y1-y0<25:continue
        nx0,nx1=name_bounds
        name_img=gray[y0+7:y1-7,nx0+6:nx1-6]
        detected=re.sub(r'\s+',' ',pytesseract.image_to_string(name_img,config='--psm 6')).strip()
        # Some names wrap two lines; OCR is only a draft and is editable in review.
        if not detected:
            detected=f'[REVIEW NAME ROW {len(candidates)+1}]'
        row=[]
        for x0,x1 in day_bounds:
            pad=max(5,int(min(x1-x0,y1-y0)*.09))
            ylo,yhi=y0+pad,y1-pad; xlo,xhi=x0+pad,x1-pad
            if yhi<=ylo or xhi<=xlo:
                row.append('unknown');continue
            region=hsv[ylo:yhi,xlo:xhi]
            rh,rw=region.shape[:2]
            # Colored ballpoint ink: substantial saturation and relatively dark.
            blue=((region[:,:,0]>=92)&(region[:,:,0]<=165)&
                  (region[:,:,1]>=35)&(region[:,:,2]<220)).astype('uint8')
            n_blue=cv2.countNonZero(blue)
            # Handle dark pen signatures independent of hue, filtering tiny artifacts.
            g=gray[ylo:yhi,xlo:xhi]
            dark=(g<115).astype('uint8')
            n_dark=cv2.countNonZero(dark)
            area=rh*rw
            # Count components to ignore small scanner dots.
            count, comps, stats, centroids=cv2.connectedComponentsWithStats(blue,8)
            significant=sum(int(stats[k,cv2.CC_STAT_AREA]) for k in range(1,count)
                            if stats[k,cv2.CC_STAT_AREA]>=4)
            if significant>=max(18,int(area*.0015)) or n_dark>=max(90,int(area*.018)):
                row.append('present')
            elif n_blue>=9 or n_dark>=max(25,int(area*.005)):
                row.append('unknown')
            else:row.append('absent')
        candidates.append((detected,row))
    if not candidates:raise ValueError('No participant rows detected in the attendance photo.')
    return candidates,labels,'Scanned attendance'

def review_image_attendance(parsed):
    """Mandatory human review: names and every day status are editable."""
    people,labels,sheet_name=parsed
    print('\nIMAGE REVIEW - OCR can misspell names and misclassify marks.')
    print('Compare each row with the original photo. Edit names/statuses below.')
    print('Status: p=present, a=absent, u=uncertain; empty input accepts suggestion.')
    labels=[input(f'Day {j+1} label (Enter={lab}): ').strip() or lab for j,lab in enumerate(labels)]
    reviewed=[]
    for i,(name,statuses) in enumerate(people,1):
        print(f'\nRow {i}/{len(people)}: {name} | '+', '.join(f'{labels[d]}={s}' for d,s in enumerate(statuses)))
        corrected=input('Name (Enter to accept): ').strip() or name
        if corrected.startswith('[REVIEW NAME'):
            raise ValueError(f'Row {i} name must be entered correctly before continuing.')
        new_status=[]
        for d,status in enumerate(statuses):
            while True:
                val=input(f'  {labels[d]} status [p/a/u] (Enter={status}): ').strip().lower()
                if not val:
                    if status=='unknown':
                        print('  Please choose p, a, or u for uncertain marks.');continue
                    new_status.append(status);break
                if val in {'p','a','u'}:
                    new_status.append({'p':'present','a':'absent','u':'unknown'}[val]);break
                print('  Please use p, a or u.')
        reviewed.append((corrected,new_status))
    print(f'\nReviewed: {len(reviewed)} names, {sum("present" in st for _,st in reviewed)} distinct attendees.')
    if input('Confirm all rows against image and generate PowerPoint? [yes/no]: ').strip().lower()!='yes':
        raise ValueError('Image attendance not confirmed; report was not generated.')
    return reviewed,labels,sheet_name

def clone_attendance_into_report(report, attend_slide):
    dest=report.slides.add_slide(report.slide_layouts[-1])
    target=dest.shapes._spTree
    for el in list(target)[2:]:target.remove(el)
    for el in list(attend_slide.shapes._spTree)[2:]:target.append(deepcopy(el))
    for child in list(dest._element.cSld):
        if child.tag==qn('p:bg'):dest._element.cSld.remove(child)
    for child in attend_slide._element.cSld:
        if child.tag==qn('p:bg'):dest._element.cSld.insert(0,deepcopy(child))
    relmap={}
    for rid,rel in attend_slide.part.rels.items():
        if rel.reltype==RT.IMAGE:
            im=rel.target_part
            name=report.part.package.next_partname('/ppt/media/image%d.'+im.partname.ext)
            new_img=ImagePart(name,im.content_type,report.part.package,im.blob)
            relmap[rid]=dest.part.relate_to(new_img,RT.IMAGE)
    for el in dest._element.iter():
        for attr in (qn('r:embed'),qn('r:link'),qn('r:id')):
            rid=el.get(attr)
            if rid in relmap:el.set(attr,relmap[rid])
    return dest

def combine_report_and_attendance(attendance_pptx, output=OUTPUT_PPTX):
    report=Presentation(REPORT_PPTX);reference=Presentation(attendance_pptx)
    if (report.slide_width,report.slide_height)!=(reference.slide_width,reference.slide_height):
        raise ValueError('Attendance and report slide sizes differ; cannot insert without distortion.')
    assert len(reference.slides)==1
    clone_attendance_into_report(report,reference.slides[0])
    ids=report.slides._sldIdLst
    entry=ids[-1];ids.remove(entry);ids.insert(3,entry)
    report.save(output)
    final=Presentation(output)
    assert len(final.slides)==len(report.slides)
    tables=[s for s in final.slides[3].shapes if s.has_table]
    assert len(tables)==2, 'Attendance slide 4 missing its editable tables.'
    print('FINAL: one striped attendance slide inserted at slide 4; all other report slides retained.')
    return output

def upload_one(prompt, ext):
    from google.colab import files
    print('\n'+prompt)
    u=files.upload()
    matches=[x for x in u if x.lower().endswith(ext)]
    if len(matches)!=1 or len(u)!=1:raise ValueError(f'Upload one {ext} file, got {list(u)}')
    return matches[0]

def run():
    from google.colab import files
    template, attendance_template = install_embedded_templates()
    survey_pdf = upload_one('STEP 1/3: Select SURVEY RESULTS (.pdf)', '.pdf')
    responses = upload_one('STEP 2/3: Select SURVEY RESPONSES (.xlsx)', '.xlsx')
    mode = input('Attendance source: 1 = Excel sheet, 2 = scanned image: ').strip().lower()
    if mode in ('1','sheet','excel'):
        attendance_file = upload_one('STEP 3/3: Select ATTENDANCE EXCEL (.xlsx)', '.xlsx')
        attendance_wb = load_workbook(attendance_file, data_only=True)
        available = [v for ws in attendance_wb.worksheets if (v := parse_sheet(ws))]
        if not available: raise ValueError('No attendance records found in Excel.')
        parsed = max(available, key=lambda x: len(x[0]))
        if pathlib.Path(responses).stem.split(' - ')[0].strip().casefold() != pathlib.Path(attendance_file).stem.split(' - ')[0].strip().casefold():
            print('WARNING: survey and attendance files might be from different workshops.')
    elif mode in ('2','image','photo'):
        attendance_file = upload_one('STEP 3/3: Select ATTENDANCE IMAGE (.jpg/.jpeg/.png)', ('.jpg','.jpeg','.png'))
        parsed = review_image_attendance(read_attendance_image(attendance_file))
    else:
        raise ValueError('Select 1 (Excel) or 2 (image).')
    build(template, survey_pdf, responses, REPORT_PPTX,
          attendance_people=parsed[0], attendance_days=len(parsed[1]))
    attendance_pptx, _ = make_attendance_one_slide(attendance_file, attendance_template, parsed=parsed)
    result = combine_report_and_attendance(attendance_pptx)
    files.download(result)
    files.download('Evaluation_Report_Automated.audit.json')

# The Gradio interface below replaces the old command-line run() call.


# ---------- SIMPLE WEB INTERFACE (Gradio) ----------
import os
import tempfile
import threading
import traceback
import pandas as pd
import gradio as gr

JOB_LOCK = threading.Lock()
TEMPLATE_CACHE = None
STATUS_OPTIONS = ('present', 'absent', 'unknown')


def get_templates_once():
    """One-time setup. No repeated PPTX uploads."""
    global TEMPLATE_CACHE
    if TEMPLATE_CACHE is not None and all(pathlib.Path(p).is_file() for p in TEMPLATE_CACHE):
        return TEMPLATE_CACHE
    local_dir = os.environ.get('EVALUATION_TEMPLATE_DIR')
    if local_dir or pathlib.Path(__file__).resolve().parent.joinpath('templates').is_dir():
        folder = pathlib.Path(local_dir) if local_dir else pathlib.Path(__file__).resolve().parent / 'templates'
        paths = (folder / REPORT_TEMPLATE_FILENAME, folder / ATTENDANCE_TEMPLATE_FILENAME)
        if not all(p.is_file() for p in paths):
            raise FileNotFoundError('Both PowerPoint templates must be placed in ' + str(folder))
        TEMPLATE_CACHE = tuple(map(str,paths))
    else:
        TEMPLATE_CACHE = install_embedded_templates()  # One Google Drive mount per runtime
    return TEMPLATE_CACHE


def attendance_to_grid(parsed):
    people, labels, _ = parsed
    cells = []
    for name, statuses in people:
        cells.append([str(name), statuses[0] if statuses else 'unknown',
                      statuses[1] if len(statuses)>1 else ''])
    return pd.DataFrame(cells, columns=['Participant name','Day 1','Day 2'])


def prepare_attendance(mode, file_path):
    """Extract a draft, not a final attendance record. Editable in the interface."""
    if not file_path:
        raise gr.Error('Please upload an attendance Excel sheet or image first.')
    mode = str(mode).lower()
    suffix = pathlib.Path(file_path).suffix.lower()
    if mode == 'excel sheet':
        if suffix not in ('.xlsx','.xlsm'):
            raise gr.Error('Excel mode requires an .xlsx or .xlsm file.')
        book=load_workbook(file_path,data_only=True)
        found=[item for ws in book.worksheets if (item:=parse_sheet(ws))]
        if not found:raise gr.Error('No attendance names and day columns were detected in this sheet.')
        parsed=max(found,key=lambda x:len(x[0]))
    else:
        if suffix not in ('.jpg','.jpeg','.png','.webp'):
            raise gr.Error('Image mode requires a JPG, PNG or WEBP photo.')
        parsed=read_attendance_image(file_path)
    people,labels,sheet=parsed
    if not 1<=len(labels)<=2:
        raise gr.Error('The template supports one or two attendance days.')
    if len(people)>76:
        raise gr.Error('More than 76 participants cannot fit on the original attendance slide.')
    note=f'**{len(people)} participants detected.** Review names, both attendance columns, and the dates before generating.'
    if mode!='excel sheet':
        note += ' **Image OCR is only a draft.** Every name and mark must be checked against the photo.'
    note += '\n\nValid statuses: `present`, `absent`, `unknown`. An empty Day 2 is allowed for one-day workshops.'
    return (attendance_to_grid(parsed),labels[0],labels[1] if len(labels)>1 else '',
            {'day_count':len(labels),'source':sheet,'mode':mode},note)


def grid_to_attendance(table, first_day, second_day, extraction):
    if not extraction:
        raise gr.Error('Click Read attendance before generating the report.')
    if not first_day or not str(first_day).strip():
        raise gr.Error('Enter the Day 1 header.')
    days=int(extraction['day_count'])
    if days==2 and not str(second_day or '').strip():
        raise gr.Error('Enter the Day 2 header.')
    if isinstance(table,pd.DataFrame):
        rows=table.fillna('').values.tolist()
    elif isinstance(table,list):
        rows=table
    else:
        raise gr.Error('Attendance table was not available. Read attendance again.')
    records=[]
    for i,row in enumerate(rows,1):
        if not row or not str(row[0]).strip():
            raise gr.Error(f'Participant name is missing in row {i}.')
        name=str(row[0]).strip()
        if name.startswith('[REVIEW NAME'):
            raise gr.Error(f'Please correct the OCR participant name in row {i}.')
        statuses=[]
        for day in range(days):
            val=str(row[day+1] if len(row)>day+1 else '').strip().lower()
            if val in ('p','yes','attended','true','1','✓','✔','✅'):val='present'
            if val in ('a','no','false','0','✗','❌'):val='absent'
            if val in ('u','uncertain',''):val='unknown'
            if val not in STATUS_OPTIONS:
                raise gr.Error(f'Row {i}, Day {day+1}: use present, absent or unknown, not {val!r}.')
            statuses.append(val)
        records.append((name,statuses))
    if not 1<=len(records)<=76:
        raise gr.Error('Attendance must contain 1 to 76 participants.')
    labels=[str(first_day).strip()] + ([str(second_day).strip()] if days==2 else [])
    return records,labels,extraction['source']


def generate_report(pdf_path, response_path, attendance_path, mode,
                    table, first_day, second_day, extraction, image_confirmed,
                    progress=gr.Progress(track_tqdm=False)):
    if not pdf_path or not response_path or not attendance_path:
        raise gr.Error('Upload the survey PDF, survey Excel and attendance file.')
    if pathlib.Path(pdf_path).suffix.lower()!='.pdf':
        raise gr.Error('The survey report must be a PDF.')
    if pathlib.Path(response_path).suffix.lower() not in ('.xlsx','.xlsm'):
        raise gr.Error('Survey responses must be an Excel workbook.')
    if not extraction:
        raise gr.Error('First click Read attendance and review the extracted rows.')
    if str(mode).lower()!='excel sheet' and not image_confirmed:
        raise gr.Error('For image attendance, confirm that you reviewed every row against the photo.')
    parsed=grid_to_attendance(table,first_day,second_day,extraction)
    if extraction['mode'] != str(mode).lower():
        raise gr.Error('Attendance mode changed: click Read attendance again.')
    report_template,attendance_template=get_templates_once()
    job_dir=pathlib.Path(tempfile.mkdtemp(prefix='evaluation_job_'))
    report_path=str(job_dir/'Evaluation_Report_Automated.pptx')
    attend_path=str(job_dir/'Attendance_Only_One_Slide_Striped.pptx')
    result_path=str(job_dir/'Evaluation_Report_With_Attendance_FINAL.pptx')
    global REPORT_PPTX
    with JOB_LOCK:  # existing merger uses REPORT_PPTX; serialize concurrent jobs
        original_report_path=REPORT_PPTX
        try:
            REPORT_PPTX=report_path
            progress(0.1,desc='Calculating survey and workshop metrics')
            build(report_template,pdf_path,response_path,report_path,
                  attendance_people=parsed[0], attendance_days=len(parsed[1]))
            progress(0.6,desc='Populating the formatted attendance slide')
            make_attendance_one_slide(attendance_path,attendance_template,output=attend_path,parsed=parsed)
            progress(0.85,desc='Combining slides and validating PowerPoint')
            combine_report_and_attendance(attend_path,output=result_path)
        finally:
            REPORT_PPTX=original_report_path
    present=sum('present' in statuses for _,statuses in parsed[0])
    msg=(f'### Report ready\n\n'
         f'**{len(parsed[0])} planned participants**, **{present} attended** '
         f'(on at least one day), **{len(parsed[1])*8} scheduled training hours**. '
         f'The attendance slide is at slide 4.\n\n'
         f'All chart styling, Appendix and comment layouts are kept from the original templates.')
    progress(1,desc='Ready')
    return result_path,str(pathlib.Path(report_path).with_suffix('.audit.json')),msg


def launch_interface():
    # Hosted app: load local bundled templates, never request Google Drive or Colab login.
    get_templates_once()
    with gr.Blocks(title='Workshop Evaluation Report Generator') as demo:
        gr.Markdown('## Workshop Evaluation Report Generator\nUpload survey data, check attendance and download the final PowerPoint. Original templates are loaded automatically.')
        with gr.Row():
            survey_pdf=gr.File(label='1 · Survey results PDF',file_types=['.pdf'],type='filepath')
            survey_xlsx=gr.File(label='2 · Survey responses Excel',file_types=['.xlsx','.xlsm'],type='filepath')
        mode=gr.Radio(['Excel sheet','Scanned image'],value='Excel sheet',label='3 · Attendance source')
        attendance_file=gr.File(label='Attendance Excel (.xlsx)',file_types=['.xlsx','.xlsm'],type='filepath')
        read_button=gr.Button('Read attendance',variant='secondary')
        extract_state=gr.State(value=None)
        notice=gr.Markdown('Upload attendance, then click **Read attendance**.')
        with gr.Accordion('Review and correct attendance',open=True):
            table=gr.Dataframe(headers=['Participant name','Day 1','Day 2'],datatype=['str','str','str'],interactive=True,
                               col_count=(3,'fixed'),row_count='dynamic',label='Editable attendance records')
            with gr.Row():
                day1=gr.Textbox(label='Day 1 date',placeholder='21 Oct')
                day2=gr.Textbox(label='Day 2 date (leave blank for one day)',placeholder='22 Oct')
            confirmed=gr.Checkbox(label='For scanned images: I reviewed every name and signature against the photo.',value=False)
        generate_button=gr.Button('Generate PowerPoint',variant='primary')
        report_status=gr.Markdown()
        with gr.Row():
            final_file=gr.File(label='Download final PowerPoint')
            audit_file=gr.File(label='Download calculation audit JSON')
        def switch_mode(selected):
            excel=selected=='Excel sheet'
            return (gr.update(label='Attendance Excel (.xlsx)' if excel else 'Attendance photo (.jpg / .png)',
                              file_types=['.xlsx','.xlsm'] if excel else ['.jpg','.jpeg','.png','.webp'],value=None),
                    None,gr.update(value=pd.DataFrame(columns=['Participant name','Day 1','Day 2'])),
                    gr.update(value=''),gr.update(value=''),gr.update(value=False),
                    'Upload attendance and click **Read attendance**.')
        mode.change(switch_mode,inputs=[mode],outputs=[attendance_file,extract_state,table,day1,day2,confirmed,notice])
        read_button.click(prepare_attendance,inputs=[mode,attendance_file],
                          outputs=[table,day1,day2,extract_state,notice])
        generate_button.click(generate_report,
             inputs=[survey_pdf,survey_xlsx,attendance_file,mode,table,day1,day2,extract_state,confirmed],
             outputs=[final_file,audit_file,report_status],show_progress='full')
    demo.queue(default_concurrency_limit=1)
    demo.launch(server_name='0.0.0.0', server_port=int(os.environ.get('PORT', '7860')), share=False, show_error=True)
    return demo

if __name__ == '__main__':
    launch_interface()
