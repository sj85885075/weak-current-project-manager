"""Export a subsystem snapshot without changing project source files."""
import base64
import html
import io
import re
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def build_bundle(data):
    tables = data.get('tables', [])
    charts = data.get('charts', [])
    if not isinstance(tables, list) or not 1 <= len(tables) <= 40:
        raise ValueError('导出表格数量无效')
    if not isinstance(charts, list) or not 1 <= len(charts) <= 20:
        raise ValueError('未取得图形，请等待图表加载后重试')
    title = str(data.get('name', '子系统'))[:100]
    safe_name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', title).strip('. ') or '子系统'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    report = ['<!doctype html><meta charset="utf-8"><title>' + html.escape(title) + '</title>',
              '<style>body{font:14px Arial,sans-serif;margin:30px;color:#243247}table{border-collapse:collapse;width:100%;margin-bottom:24px}th,td{border:1px solid #ccd4df;padding:8px;text-align:left;white-space:pre-wrap}th{background:#e8eef6}img{width:100%;height:auto}h2{break-after:avoid}@media print{img,tr{break-inside:avoid}}</style>',
              '<h1>' + html.escape(title) + '：全部表格与图形</h1>',
              '<p>填报日期：' + html.escape(str(data.get('date', ''))) + '；统计周期：' + html.escape(str(data.get('period', ''))) + '</p>',
              '<p>本文件为点击导出时的页面快照，包含当前输入值，不会自动保存到项目台账。图形及计量流水采用已保存数据；图形展示完整时间范围。乐观/悲观按累计计划120%/80%计算，上限为计划总量。</p>']
    images = []
    for i, table in enumerate(tables, 1):
        name = str(table.get('name', '表格'))[:100]
        rows = table.get('rows', [])
        if not isinstance(rows, list) or len(rows) > 50000:
            raise ValueError('表格行数无效')
        sheet_name = re.sub(r'[\\/*?:\[\]]', '_', f'{i:02d}_{name}')[:31]
        ws = wb.create_sheet(sheet_name)
        report.append('<h2>' + html.escape(name) + '</h2><table>')
        for n, row in enumerate(rows):
            if not isinstance(row, list) or len(row) > 100:
                raise ValueError('表格列数无效')
            for c, value in enumerate(row, 1):
                value = value if isinstance(value, (int, float)) else str(value if value is not None else '')[:32767]
                cell = ws.cell(n + 1, c, value)
                if isinstance(value, str):
                    cell.data_type = 's'  # User-entered text must never become an Excel formula.
                cell.alignment = Alignment(vertical='top', wrap_text=True)
            tag = 'th' if n == 0 else 'td'
            report.append('<tr>' + ''.join(f'<{tag}>{html.escape(str(v if v is not None else ""))}</{tag}>' for v in row) + '</tr>')
        report.append('</table>')
        for cell in ws[1]:
            cell.font = Font(name='微软雅黑', color='FFFFFF', bold=True)
            cell.fill = PatternFill('solid', fgColor='243B60')
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = min(48, max(16, max((len(str(ws.cell(r, col).value or '')) for r in range(1, min(ws.max_row, 100) + 1)), default=10) * 1.4))
    for i, chart in enumerate(charts, 1):
        name = str(chart.get('name', '图形'))[:100]
        encoded = str(chart.get('image', ''))
        if not encoded.startswith('data:image/png;base64,'):
            raise ValueError('图形必须为 PNG')
        image = base64.b64decode(encoded.split(',', 1)[1], validate=True)
        if not image.startswith(b'\x89PNG\r\n\x1a\n') or len(image) > 12 * 1024 * 1024:
            raise ValueError('PNG 图形无效或过大')
        filename = f'图形/{i:02d}_' + re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', name) + '.png'
        images.append((filename, image))
        report.append('<h2>' + html.escape(name) + '</h2><img alt="' + html.escape(name, quote=True) + '" src="' + html.escape(encoded, quote=True) + '">')
    excel = io.BytesIO()
    wb.save(excel)
    result = io.BytesIO()
    with zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('全部表格.xlsx', excel.getvalue())
        archive.writestr('完整报告.html', '\n'.join(report).encode('utf-8'))
        for name, image in images:
            archive.writestr(name, image)
    return result.getvalue(), safe_name + '_全部表格与图形.zip'
