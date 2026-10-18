"""Editable Word owner report, using the same schedule and charts as overview."""
import io
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from PIL import Image

import construction_flow


def chart_images(model, title):
    """Rasterize the exact SVG charts from the report with bundled sharp."""
    markup = construction_flow.report(model, title)
    charts = re.findall(r'<svg\b.*?</svg>', markup, flags=re.S)
    dependencies = Path(sys.executable).resolve().parent.parent
    if getattr(sys, 'frozen', False):
        dependencies = Path(sys.executable).resolve().parent / 'runtime'
    node = dependencies / 'node' / 'bin' / 'node.exe'
    sharp = dependencies / 'node' / 'node_modules' / 'sharp'
    if not node.exists() or not sharp.exists():
        raise RuntimeError('Word 图表导出需要启动工具配置的完整运行环境，请使用启动工具.bat')
    script = "const fs=require('fs');const sharp=require(process.argv[1]);sharp(fs.readFileSync(process.argv[2]),{density:216}).flatten({background:'#ffffff'}).png().toFile(process.argv[3]).catch(e=>{console.error(e);process.exit(1)});"
    result = []
    with tempfile.TemporaryDirectory(prefix='construction_word_') as folder:
        for i, svg in enumerate(charts):
            svg = svg.replace('<svg ', '<svg font-family="Microsoft YaHei, sans-serif" ', 1)
            source, target = Path(folder)/f'chart{i}.svg', Path(folder)/f'chart{i}.png'
            source.write_text(svg, encoding='utf-8')
            subprocess.run([str(node), '-e', script, str(sharp), str(source), str(target)],
                           check=True, capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            result.append(target.read_bytes())
    if len(result) != 2:
        raise ValueError('施工汇报图表不完整')
    return result


def make_report(model, title):
    images = chart_images(model, title)
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Mm(297), Mm(210)
    section.top_margin = section.bottom_margin = Mm(15)
    section.left_margin = section.right_margin = Mm(17)
    for name in ('Normal', 'Title', 'Heading 1', 'Heading 2', 'Caption'):
        style = doc.styles[name]
        style.font.name = 'Microsoft YaHei'
        style._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), '微软雅黑')
        style.font.color.rgb = RGBColor(0, 0, 0)
    normal = doc.styles['Normal']
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    doc.styles['Title'].font.size = Pt(23)
    doc.styles['Title'].paragraph_format.space_after = Pt(8)
    doc.styles['Heading 1'].font.size = Pt(16)
    doc.styles['Heading 1'].paragraph_format.space_before = Pt(10)
    doc.styles['Heading 1'].paragraph_format.space_after = Pt(7)
    doc.core_properties.title = title + '施工进度汇报'
    doc.core_properties.subject = '子系统施工流程 工期安排及并行关系'
    doc.core_properties.author = '项目管理组'
    doc.add_paragraph(title + '施工进度汇报', 'Title')
    doc.add_paragraph('数据时间：' + model['generated'].replace('T', ' ') + '    汇报对象：甲方项目管理团队')
    total = (str(model['total'])+' 个自然日') if model['complete'] else ('待确定，已排部分跨度为 '+str(model['total'])+' 个自然日')
    doc.add_paragraph('本报告汇总子系统施工安排、先后关系及并行区间，供甲方核对总体进度和协调施工条件。当前同步 '+str(len(model['nodes']))+' 个子系统工作包，总工期为 '+total+'。')

    def table(headers, rows, widths):
        t = doc.add_table(rows=1, cols=len(headers))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for col, width in zip(t.columns, widths): col.width = Mm(width)
        for i, h in enumerate(headers): t.rows[0].cells[i].text = str(h)
        repeat = OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(repeat)
        for row in rows:
            for cell, value in zip(t.add_row().cells, row): cell.text = str(value)
        for ri, row in enumerate(t.rows):
            no_split = OxmlElement('w:cantSplit');row._tr.get_or_add_trPr().append(no_split)
            for ci, cell in enumerate(row.cells):
                cell.width = Mm(widths[ci]); cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                props = cell._tc.get_or_add_tcPr()
                shade = OxmlElement('w:shd');shade.set(qn('w:fill'), 'DBEAFE' if ri == 0 else 'FFFFFF');props.append(shade)
                borders = OxmlElement('w:tcBorders')
                for side in ('top','left','bottom','right'):
                    border = OxmlElement('w:'+side)
                    for k,v in {'val':'single','sz':'4','color':'D9D9D9'}.items():border.set(qn('w:'+k),v)
                    borders.append(border)
                props.append(borders)
                margins = OxmlElement('w:tcMar')
                for side in ('top','left','bottom','right'):
                    margin = OxmlElement('w:'+side);margin.set(qn('w:w'),'85');margin.set(qn('w:type'),'dxa');margins.append(margin)
                props.append(margins)
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(2);p.paragraph_format.space_before = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci in (0,1) else WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:r.font.size=Pt(9);r.bold=(ri==0)
        doc.add_paragraph().paragraph_format.space_after=Pt(0)
        return t

    table(['计划开始','排程完成' if model['complete'] else '已排部分完成','总工期','可并行组合'],
          [[model['start'] or '相对排程',model['finish'] or '待定',total,len(model['parallel'])]], [60,60,103,40])
    if model['warnings']:
        doc.add_heading('待确认事项', 1)
        for warning in model['warnings']:doc.add_paragraph(warning, style='List Bullet')
    doc.add_paragraph('工期口径：按自然日计算，含开工和完工当天。工作包开始日期为最早开工约束；未校核人员、材料、场地资源及节假日。无明确依赖时，仅依据日期识别先后和重叠施工区间。')

    def chart(data, caption, max_height=135):
        # Keep the complete diagram together; editable tables retain all labels.
        image = Image.open(io.BytesIO(data)).convert('RGB')
        width=min(263, max_height*image.width/image.height)
        output=io.BytesIO();image.save(output,format='PNG');output.seek(0)
        p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        picture=p.add_run().add_picture(output,width=Mm(width))
        picture._inline.docPr.set('descr',caption)
        doc.add_paragraph(caption,'Caption')

    doc.add_page_break();doc.add_heading('施工流程与工期关系',1)
    doc.add_paragraph('横向按前置层级排列；未配置依赖时按开始日期排列。同日开始的工作包并列显示。箭头标注关系类型和间隔天数，红框表示当前约束下零总时差工作。')
    chart(images[0],'图1 子系统施工流程图')
    doc.add_page_break();doc.add_heading('施工进度甘特图',1)
    doc.add_paragraph('横轴 D+ 表示从计划开始日起经过的自然日。浅色条为排程工期，深色部分为实际完成比例；红色为零总时差工作。')
    chart(images[1],'图2 子系统工期与实际进度')
    doc.add_page_break();doc.add_heading('子系统工作包排程明细',1)
    def dates(n):
        if n['start'] is None:return '待排程'
        return (n['start_date'] or '第%d天'%(n['start']+1))+'\n至 '+(n['finish_date'] or '第%d天'%n['end'])
    table(['WBS','子系统工作包','责任人','自然日','排程起止','完成率','总时差','工期来源'],
          [[n['code'],n['name'],n['owner'],n['duration'] or '待定',dates(n),str(n['progress'])+'%' if n['progress'] is not None else '无计量基数',
            '关键工作' if n['float']==0 else (str(n['float'])+'天' if n['float'] is not None else '待定'),'单独设置' if n['override'] else '工作包同步'] for n in model['nodes']],
          [19,68,30,18,37,25,28,38])
    doc.add_heading('前后置关系',1)
    names={n['id']:n['name'] for n in model['nodes']}
    meanings={'FS':'完成后开始','SS':'开始后开始','FF':'完成后完成'}
    if model['edges']:
        table(['前置子系统','后续子系统','工期关系','间隔自然日'],[[names[e['source']],names[e['target']],e['type']+' '+meanings[e['type']],e['lag']] for e in model['edges']],[85,85,60,33])
    else:doc.add_paragraph('目前未配置跨子系统前置关系。请确认施工依赖后重新导出正式排程。')
    doc.add_paragraph('FS 0天：前项完成后次日开始；SS 0天：后项可与前项同日开始；FF 0天：后项完成时间不得早于前项。正间隔为滞后，负间隔为提前。')
    doc.add_heading('可并行施工区间',1)
    by={n['id']:n for n in model['nodes']}
    overlaps=[]
    for e in model['parallel']:
        a,b=by[e['source']],by[e['target']]
        later=max(a['start'],b['start']);earlier=min(a['end'],b['end'])
        import datetime
        base=construction_flow.day(model['start'])
        span=(str(base+datetime.timedelta(days=later))+' 至 '+str(base+datetime.timedelta(days=earlier-1))) if base else f'第{later+1}天 至 第{earlier}天'
        overlaps.append([names[e['source']],names[e['target']],span,str(e['days'])+'天'])
    if overlaps:table(['子系统一','子系统二','重叠日期','重叠工期'],overlaps,[82,82,72,27])
    else:doc.add_paragraph('当前已排工作包没有日期重叠的并行组合。')
    doc.add_paragraph('并行区间按排程日期交集计算，实际同时施工需落实班组、材料、作业面及安全条件。')
    buffer=io.BytesIO();doc.save(buffer)
    return buffer.getvalue()
