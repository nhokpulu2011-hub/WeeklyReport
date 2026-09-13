import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, date, timedelta
from pathlib import Path

import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    from tkcalendar import DateEntry
    HAVE_TKCALENDAR = True
except ImportError:
    HAVE_TKCALENDAR = False

# ================================================================
# SY1/2 WEEKLY REPORT GENERATOR
# Rebuilds the report from scratch using python-docx.
# Weekly stats + Customer/Internal project updates come from an
# Excel workbook (read with openpyxl) keyed on "Week Start" date.
# ================================================================

A3_WIDTH = Inches(11.69)
A3_HEIGHT = Inches(16.54)
MARGIN_LR = Inches(1.00)
MARGIN_TB = Inches(0.60)
CONTENT_W = Inches(7.63)
FONT = "Calibri"

PINK = "EDBAB1"
GREY = "C4C4D0"
GREEN = "D9EAD3"
DARK_GREEN = "70AD47"
PALE_YELLOW = "FFF2CC"
PALE_RED = "F4CCCC"
NAVY = "002060"
DARK = "292934"
RED = "FF0000"
WHITE = "FFFFFF"
BLACK = "000000"


# ----------------------------------------------------------------
# docx low-level helpers (unchanged from the original)
# ----------------------------------------------------------------

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:fill'), fill)
    shd.set(qn('w:val'), 'clear')


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge in kwargs:
            edge_data = kwargs.get(edge)
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ['val', 'sz', 'space', 'color']:
                if key in edge_data:
                    element.set(qn('w:' + key), str(edge_data[key]))


def set_table_borders(table, color='A6A6A6', size='4'):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell,
                            top={'val': 'single', 'sz': size, 'color': color},
                            bottom={'val': 'single', 'sz': size, 'color': color},
                            left={'val': 'single', 'sz': size, 'color': color},
                            right={'val': 'single', 'sz': size, 'color': color})


def set_cell_margins(cell, top=60, start=70, bottom=60, end=70):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in('w:tcMar')
    if tcMar is None:
        tcMar = OxmlElement('w:tcMar')
        tcPr.append(tcMar)
    for m, v in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tcMar.find(qn(f'w:{m}'))
        if node is None:
            node = OxmlElement(f'w:{m}')
            tcMar.append(node)
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'), 'dxa')


def set_row_height(row, twips, exact=True):
    trPr = row._tr.get_or_add_trPr()
    trHeight = trPr.find(qn('w:trHeight'))
    if trHeight is None:
        trHeight = OxmlElement('w:trHeight')
        trPr.append(trHeight)
    trHeight.set(qn('w:val'), str(twips))
    trHeight.set(qn('w:hRule'), 'exact' if exact else 'atLeast')


def set_cell_width(cell, width_inches):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = OxmlElement('w:tcW')
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(width_inches * 1440)))
    tcW.set(qn('w:type'), 'dxa')


def set_table_width(table, width_inches):
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:w'), str(int(width_inches * 1440)))
    tblW.set(qn('w:type'), 'dxa')


def set_fixed_layout(table):
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')


def no_table_autofit(table):
    table.autofit = False
    set_fixed_layout(table)


def style_run(run, size=10, bold=False, color=BLACK, highlight=None,
              underline=False, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline
    run.font.color.rgb = RGBColor.from_string(color)
    if highlight:
        run.font.highlight_color = getattr(
            __import__('docx.enum.text', fromlist=['WD_COLOR_INDEX']).WD_COLOR_INDEX, highlight)


def clear_cell(cell):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    return p


def add_text(p, text, size=10, bold=False, color=BLACK, highlight=None,
             underline=False, font=FONT):
    r = p.add_run(text)
    style_run(r, size=size, bold=bold, color=color, highlight=highlight,
              underline=underline, font=font)
    return r


def set_no_split(row):
    trPr = row._tr.get_or_add_trPr()
    cantSplit = OxmlElement('w:cantSplit')
    trPr.append(cantSplit)


def page_setup(doc):
    sec = doc.sections[0]
    sec.page_width = A3_WIDTH
    sec.page_height = A3_HEIGHT
    sec.top_margin = MARGIN_TB
    sec.bottom_margin = MARGIN_TB
    sec.left_margin = MARGIN_LR
    sec.right_margin = MARGIN_LR
    sec.header_distance = Inches(0.25)
    sec.footer_distance = Inches(0.18)


def style_document(doc):
    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.line_spacing = 1.0


def add_outer_table(doc, rows=1):
    tbl = doc.add_table(rows=rows, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_autofit(tbl)
    set_table_width(tbl, CONTENT_W.inches)
    for row in tbl.rows:
        set_no_split(row)
        set_cell_width(row.cells[0], CONTENT_W.inches)
        row.cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        set_cell_margins(row.cells[0], top=0, start=80, bottom=0, end=80)
    set_table_borders(tbl, color='666666', size='6')
    return tbl


def title_block(cell, title):
    p = clear_cell(cell)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    add_text(p, title, size=14, bold=True, color=BLACK)
    set_cell_shading(cell, PINK)
    set_cell_margins(cell, top=70, start=60, bottom=70, end=60)


def section_bar(cell, text, fill=GREY):
    p = clear_cell(cell)
    add_text(p, text, size=9, bold=True, color=DARK)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, top=45, start=70, bottom=45, end=70)


def add_instructions(cell):
    p = clear_cell(cell)
    p.paragraph_format.space_after = Pt(1)
    add_text(p, 'Instructions:', size=10, bold=True, color=DARK, highlight='YELLOW')
    lines = [
        '1:  Update this document, with important events during week, Projects completed, great customer feedback, team events anything that you think upper management should know about relevant to your IBX.',
        '2: In this document update Site Name and week beginning Date above and then send to FE manager',
        '3: Make Subject of email "SY1/SY2 Weekly Report Week beginning (Date) "',
        '4: Send to FE manager by 10am Monday',
    ]
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        add_text(p, line, size=10, bold=True, color=DARK, highlight='YELLOW')

    p = cell.add_paragraph(); p.paragraph_format.space_after = Pt(1)
    add_text(p, 'NOTE: ', size=10, bold=True, color=DARK)

    notes = [
        '1: I use this report to prepare a weekly and then Monthly report for Harry, Harry uses my report to update the Senior Leadership team Monday afternoon.',
        'Your input into this report is a great time to make the Senior Leadership team aware of your team\u2019s achievements so please put some serious thought into this weekly report.',
        '2: There will be a reminder sent on Friday before and Monday.',
    ]
    for i, line in enumerate(notes):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3 if i < 2 else 5)
        p.paragraph_format.line_spacing = 1.0
        add_text(p, line, size=10, bold=True, color=DARK)

    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    add_text(p, '-' * 110, size=8, bold=False, color=DARK)


def add_stats_table(parent_cell, week_no, stats):
    """Renders items 1-8 (stats & figures + HR). Customer/Internal project
    sections are now handled separately by add_project_section() so their
    length can vary freely with however many rows the Excel sheet has."""
    p = parent_cell.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(0)
    add_text(p, 'Weekly Recap \u2013 SY1/2:', size=10, bold=True, color=DARK)

    tbl = parent_cell.add_table(rows=2, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_autofit(tbl)
    set_table_width(tbl, CONTENT_W.inches - 0.10)
    for r in tbl.rows:
        set_cell_width(r.cells[0], CONTENT_W.inches - 0.10)
        set_cell_shading(r.cells[0], GREEN)
        set_cell_margins(r.cells[0], top=45, start=70, bottom=45, end=70)
        set_no_split(r)
    set_table_borders(tbl, color='D0D0D0', size='3')

    # Stats row
    c = tbl.cell(0, 0); p = clear_cell(c)
    add_text(p, f'Stats & Figures for Week {week_no}:', size=10, bold=True, color=DARK, highlight='YELLOW')
    stat_texts = [
        (stats['smart_hands'], ' Smart Hands for the week, ', stats['smart_hands_hours'], ' hours billed.'),
        (stats['trouble_tickets'], ' Trouble ticket for the week, ', stats['trouble_ticket_hours'], ' hours billed.'),
        (stats['cross_connect_installed'], ' Cross connects installed.', None, None),
        (stats['cross_connect_deinstalled'], ' Cross connects de-installed.', None, None),
        (stats['cross_connect_migrated'], ' Cross connects migrated.', None, None),
        (stats['commit_compliance'], ' % SY1/2 Commit Compliance', None, None),
        (stats['labour_hours'], ' hrs | ', stats['labour_capitalization'] + '%', ' Labour Capitalization (SY1&SY2)'),
    ]
    for idx, (a, b, cval, d) in enumerate(stat_texts, 1):
        pp = c.add_paragraph()
        pp.paragraph_format.space_before = Pt(0)
        pp.paragraph_format.space_after = Pt(0)
        pp.paragraph_format.left_indent = Inches(0.45)
        add_text(pp, f'{idx}.    ', size=10, bold=False, color=DARK, highlight='YELLOW')
        add_text(pp, a, size=10, bold=True, color=RED, highlight='YELLOW')
        add_text(pp, b, size=10, bold=False, color=DARK, highlight='YELLOW')
        if cval is not None:
            add_text(pp, cval, size=10, bold=True, color=RED, highlight='YELLOW')
            add_text(pp, d, size=10, bold=False, color=DARK, highlight='YELLOW')

    # HR row
    c = tbl.cell(1, 0); p = clear_cell(c)
    p.paragraph_format.left_indent = Inches(0.45)
    add_text(p, '8.    ', size=10, color=DARK)
    add_text(p, 'Human Resources', size=10, bold=True, color=DARK)
    pp = c.add_paragraph(); pp.paragraph_format.left_indent = Inches(0.68)
    add_text(pp, stats.get('hr_notes') or '-', size=10, color=DARK)

    return tbl


def add_project_section(parent_cell, heading, projects):
    """Generic, dynamically-sized Customer/Internal project block.
    `projects` is a list of dicts as returned by read_customer_projects()
    / read_internal_projects() -- see those functions for the shape."""
    p = parent_cell.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    add_text(p, heading, size=10, bold=True, color=NAVY, highlight='YELLOW', underline=True)

    tbl = parent_cell.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_autofit(tbl)
    set_table_width(tbl, CONTENT_W.inches - 0.10)
    c = tbl.cell(0, 0)
    set_cell_shading(c, GREEN)
    set_cell_margins(c, top=60, start=70, bottom=60, end=70)
    set_table_borders(tbl, color='D0D0D0', size='3')
    clear_cell(c)

    if not projects:
        add_text(c.paragraphs[0], 'No items reported for this week.', size=10, color=DARK)
        return tbl

    first_para = True
    for i, proj in enumerate(projects, 1):
        p = c.paragraphs[0] if first_para else c.add_paragraph()
        first_para = False
        p.paragraph_format.left_indent = Inches(0.48)
        p.paragraph_format.space_before = Pt(6) if i > 1 else Pt(0)
        add_text(p, f'{i}.    ', size=10, color=DARK)
        add_text(p, proj['name'], size=10, bold=True, color=DARK)

        if proj.get('progress'):
            pp = c.add_paragraph(); pp.paragraph_format.left_indent = Inches(0.92)
            add_text(pp, proj['progress'], size=10, bold=True, color=RED)

        if proj.get('reference'):
            pp = c.add_paragraph(); pp.paragraph_format.left_indent = Inches(0.92)
            add_text(pp, f"Quote# {proj['reference']}", size=10, color=DARK)

        if proj.get('status'):
            pp = c.add_paragraph(); pp.paragraph_format.left_indent = Inches(0.92)
            add_text(pp, proj['status'], size=10, color=DARK)

        for line in proj.get('details', []):
            pp = c.add_paragraph(); pp.paragraph_format.left_indent = Inches(0.92)
            add_text(pp, line, size=10, color=DARK)

    return tbl


def add_shipment_table(parent_cell, shipments=None):
    p = parent_cell.add_paragraph(); p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(6)
    add_text(p, 'SHIPMENTS', size=10, bold=True, color=NAVY, underline=True)

    tbl = parent_cell.add_table(rows=3, cols=6)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_autofit(tbl)
    widths = [0.93, 0.77, 1.02, 1.10, 1.20, 1.88]
    total = sum(widths)
    set_table_width(tbl, total)
    set_cell_margins(tbl.cell(0, 0), top=55, start=30, bottom=55, end=30)

    header = tbl.cell(0, 0)
    for j in range(1, 6):
        header = header.merge(tbl.cell(0, j))
    set_cell_shading(header, DARK_GREEN)
    p = clear_cell(header); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(p, 'SY1/2 Shipment', size=14, bold=True, color=WHITE)
    set_row_height(tbl.rows[0], 390, exact=True)

    labels = ['Inbound\nShipment', 'Outbound\nShipment', 'Inbound Cage\nDelivery',
              'Actual Shipments\nReceived per\nweek', 'Shipment\nHandover per week',
              'Siebel Outstanding Shipment\nActivities']
    defaults = ['43', '19', '8', '10', '11', '79']
    values = shipments if shipments else defaults
    for j, w in enumerate(widths):
        set_cell_width(tbl.cell(1, j), w); set_cell_width(tbl.cell(2, j), w)
        set_cell_shading(tbl.cell(1, j), DARK_GREEN)
        set_cell_margins(tbl.cell(1, j), top=50, start=20, bottom=50, end=20)
        p = clear_cell(tbl.cell(1, j)); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for k, line in enumerate(labels[j].split('\n')):
            if k:
                p.add_run().add_break()
            add_text(p, line, size=9, bold=True, color=WHITE)
        p = clear_cell(tbl.cell(2, j)); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, str(values[j]), size=10, color=BLACK)
        set_cell_margins(tbl.cell(2, j), top=30, start=20, bottom=30, end=20)
        if j == 5:
            set_cell_shading(tbl.cell(2, j), DARK_GREEN)
            for r in tbl.cell(2, j).paragraphs[0].runs:
                r.font.color.rgb = RGBColor.from_string(WHITE)
                r.font.bold = True
    set_row_height(tbl.rows[1], 1515, exact=True)
    set_row_height(tbl.rows[2], 300, exact=True)
    set_table_borders(tbl, color=BLACK, size='7')
    return tbl


def add_action_box(parent_cell, actions=None):
    for _ in range(3):
        parent_cell.add_paragraph()
    tbl = parent_cell.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_autofit(tbl); set_table_width(tbl, CONTENT_W.inches - 0.20)
    c = tbl.cell(0, 0); set_cell_shading(c, PALE_YELLOW); set_cell_margins(c, top=65, start=70, bottom=65, end=70)
    if actions is None:
        actions = [
            'Be alert of COVID-19 Coronavirus to ensure the LOCC team stays healthy and safe',
            'Continue reminding team to add labor capitalization charges',
            'Continue working in maintaining the Zero OPS NCC errors in SY1/2',
            'Continue to remind the team to focus on Commit Compliance, achieving 95% target',
        ]
    clear_cell(c)
    for action in actions:
        p = c.paragraphs[0] if not c.paragraphs[0].text else c.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.12)
        p.paragraph_format.first_line_indent = Inches(-0.07)
        p.paragraph_format.space_after = Pt(0)
        add_text(p, '-    ', size=9, color=BLACK)
        add_text(p, action, size=9, color=BLACK)
    set_row_height(tbl.rows[0], 1392, exact=True)
    set_table_borders(tbl, color='D9D9D9', size='3')
    return tbl


def add_red_flags(parent_cell):
    for _ in range(3):
        parent_cell.add_paragraph()
    p = parent_cell.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    add_text(p, 'Red Flags', size=10, bold=True, color=RED)
    tbl = parent_cell.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_autofit(tbl); set_table_width(tbl, CONTENT_W.inches - 0.10)
    c = tbl.cell(0, 0); set_cell_shading(c, PALE_RED); clear_cell(c); set_cell_margins(c, top=25, start=25, bottom=25, end=25)
    set_row_height(tbl.rows[0], 323, exact=True)
    set_table_borders(tbl, color='E0BABA', size='2')


# ----------------------------------------------------------------
# Excel reading (openpyxl) -- replaces the old artifact_tool /
# SpreadsheetFile / Blob based implementation.
#
# Workbook layout (see the generated template for a filled example):
#   Sheet "Stats"             - one row per week of numeric metrics
#   Sheet "CustomerProjects"  - one row PER PROJECT ITEM, tagged with
#                                the Week Start date it belongs to
#   Sheet "InternalProjects"  - same idea, plus an optional "Progress"
#                                column (e.g. "3 of 7 complete")
#
# All three sheets are matched to the report being generated purely
# by the "Week Start" date column, so adding/removing project rows
# never requires touching the code.
# ----------------------------------------------------------------

STATS_HEADERS = [
    'Week Start', 'Smart Hands', 'Smart Hands Hours', 'Trouble Tickets',
    'Trouble Ticket Hours', 'Cross Connect Installed',
    'Cross Connect De-installed', 'Cross Connect Migrated',
    'Commit Compliance', 'Labour Hours', 'Labour Capitalization',
]
CUSTOMER_PROJECT_HEADERS = ['Week Start', 'Project Name', 'Reference', 'Status', 'Details']
INTERNAL_PROJECT_HEADERS = ['Week Start', 'Project Name', 'Reference', 'Progress', 'Status', 'Details']


def _norm(value):
    return ' '.join(str(value).strip().lower().replace('_', ' ').replace('-', ' ').split())


def _num(value):
    if value is None or value == '':
        return '0'
    try:
        x = float(value)
        return str(int(x)) if x.is_integer() else f'{x:.2f}'.rstrip('0').rstrip('.')
    except (TypeError, ValueError):
        return str(value).strip()


def _pct(value):
    if value is None or value == '':
        return '0.00'
    try:
        x = float(value)
        if 0 <= x <= 1:
            x *= 100
        return f'{x:.2f}'
    except (TypeError, ValueError):
        return str(value).strip().replace('%', '')


def _to_date(value):
    """Normalises an Excel cell value (datetime, date, or text) to a
    plain date for comparison against the week the user picked."""
    if value is None or value == '':
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _header_map(sheet):
    """Maps normalised header text -> 1-based column index, from row 1."""
    headers = {}
    for cell in sheet[1]:
        if cell.value not in (None, ''):
            headers[_norm(cell.value)] = cell.column
    return headers


def _require_headers(header_map, required, sheet_name):
    missing = [h for h in required if _norm(h) not in header_map]
    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing column(s):\n" + '\n'.join(missing)
        )


def _load_sheet(excel_path, sheet_name, required=True):
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        if required:
            raise ValueError(
                f"Workbook has no '{sheet_name}' sheet. "
                f"Available sheets: {', '.join(wb.sheetnames)}"
            )
        return None
    return wb[sheet_name]


def read_stats_from_excel(excel_path, week_start, sheet_name='Stats'):
    """Returns the stats dict for the row whose 'Week Start' matches
    `week_start` (a datetime.date). Raises ValueError if not found."""
    sheet = _load_sheet(excel_path, sheet_name)
    hmap = _header_map(sheet)
    _require_headers(hmap, STATS_HEADERS, sheet_name)

    for row in sheet.iter_rows(min_row=2):
        values = [c.value for c in row]
        if all(v in (None, '') for v in values):
            continue

        def g(name):
            idx = hmap[_norm(name)] - 1
            return values[idx] if idx < len(values) else None

        if _to_date(g('Week Start')) != week_start:
            continue
        return {
            'smart_hands': _num(g('Smart Hands')),
            'smart_hands_hours': _num(g('Smart Hands Hours')),
            'trouble_tickets': _num(g('Trouble Tickets')),
            'trouble_ticket_hours': _num(g('Trouble Ticket Hours')),
            'cross_connect_installed': _num(g('Cross Connect Installed')),
            'cross_connect_deinstalled': _num(g('Cross Connect De-installed')),
            'cross_connect_migrated': _num(g('Cross Connect Migrated')),
            'commit_compliance': _pct(g('Commit Compliance')),
            'labour_hours': _num(g('Labour Hours')),
            'labour_capitalization': _pct(g('Labour Capitalization')),
            'hr_notes': (str(g('HR Notes')).strip() if 'hr notes' in hmap and g('HR Notes') not in (None, '') else None),
        }
    raise ValueError(
        f"No row in '{sheet_name}' has Week Start = {week_start:%d/%m/%Y}."
    )


def _read_projects(excel_path, week_start, sheet_name, required_headers, has_progress):
    sheet = _load_sheet(excel_path, sheet_name, required=False)
    if sheet is None:
        return []  # project sheets are optional; an empty result is fine

    hmap = _header_map(sheet)
    _require_headers(hmap, required_headers, sheet_name)

    projects = []
    for row in sheet.iter_rows(min_row=2):
        values = [c.value for c in row]
        if all(v in (None, '') for v in values):
            continue

        def g(name):
            idx = hmap[_norm(name)] - 1
            return values[idx] if idx < len(values) else None

        if _to_date(g('Week Start')) != week_start:
            continue

        name = str(g('Project Name') or '').strip()
        if not name:
            continue  # skip stray/blank rows

        details_raw = g('Details') or ''
        details = [ln.strip() for ln in str(details_raw).splitlines() if ln.strip()]

        proj = {
            'name': name,
            'reference': str(g('Reference') or '').strip(),
            'status': str(g('Status') or '').strip(),
            'details': details,
        }
        if has_progress:
            proj['progress'] = str(g('Progress') or '').strip()
        projects.append(proj)
    return projects


def read_customer_projects(excel_path, week_start, sheet_name='CustomerProjects'):
    """One list item per Excel row tagged with this Week Start. Each row:
    Project Name | Reference | Status | Details (multi-line, Alt+Enter
    inside the cell -> each line becomes its own indented line)."""
    return _read_projects(excel_path, week_start, sheet_name,
                           CUSTOMER_PROJECT_HEADERS, has_progress=False)


def read_internal_projects(excel_path, week_start, sheet_name='InternalProjects'):
    """Same as read_customer_projects but with an extra optional
    'Progress' column (e.g. '3 of 7 complete'), shown in bold red under
    the project name."""
    return _read_projects(excel_path, week_start, sheet_name,
                           INTERNAL_PROJECT_HEADERS, has_progress=True)


# ----------------------------------------------------------------
# Report assembly
# ----------------------------------------------------------------

def build_report(output_path, week_start, week_end, stats,
                  customer_projects=None, internal_projects=None,
                  shipments=None, actions=None):
    customer_projects = customer_projects or []
    internal_projects = internal_projects or []
    week_no = week_start.isocalendar()[1]

    doc = Document()
    page_setup(doc)
    style_document(doc)

    first = add_outer_table(doc, rows=3)

    # Row 0: title
    c = first.cell(0, 0)
    title_block(c, f'(SY1/2) Weekly Report \u2013 {week_start:%d/%m/%Y} to {week_end:%d/%m/%Y}')
    set_row_height(first.rows[0], 1086, exact=True)

    # Row 1: LOCC TEAM
    section_bar(first.cell(1, 0), 'LOCC TEAM', GREY)
    set_row_height(first.rows[1], 479, exact=True)

    # Row 2: main body
    main = first.cell(2, 0)
    clear_cell(main)
    add_instructions(main)
    add_stats_table(main, week_no, stats)
    add_project_section(main, 'CUSTOMER PROJECT', customer_projects)
    add_project_section(main, 'INTERNAL PROJECT', internal_projects)

    main.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    add_shipment_table(main, shipments)
    add_action_box(main, actions)
    add_red_flags(main)

    first.add_row()
    set_cell_width(first.rows[-1].cells[0], CONTENT_W.inches)
    clear_cell(first.rows[-1].cells[0])
    set_row_height(first.rows[-1], 150, exact=True)
    set_cell_margins(first.rows[-1].cells[0], 0, 0, 0, 0)

    for p in doc.paragraphs:
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)

    doc.save(output_path)
    return output_path


# ----------------------------------------------------------------
# GUI
# ----------------------------------------------------------------

def _make_date_picker(parent, initial):
    """Returns (widget, get_date_fn). Uses tkcalendar's DateEntry when
    available for a real calendar pop-up; otherwise falls back to a
    plain dd/mm/yyyy text entry."""
    if HAVE_TKCALENDAR:
        widget = DateEntry(parent, date_pattern='dd/mm/yyyy', width=12)
        widget.set_date(initial)
        return widget, (lambda w=widget: w.get_date())
    else:
        var = tk.StringVar(value=initial.strftime('%d/%m/%Y'))
        widget = tk.Entry(parent, textvariable=var, width=14)

        def get_date(v=var):
            return datetime.strptime(v.get().strip(), '%d/%m/%Y').date()
        return widget, get_date


def gui():
    root = tk.Tk()
    root.title('SY1/2 Weekly Report Generator')
    root.geometry('760x560')
    root.resizable(False, False)

    excel_var = tk.StringVar()
    today = date.today()
    default_start = today - timedelta(days=today.weekday())  # this week's Monday

    frm = tk.Frame(root, padx=18, pady=18)
    frm.pack(fill='both', expand=True)
    tk.Label(frm, text='SY1/2 Weekly Report Generator', font=('Segoe UI', 16, 'bold')).pack(anchor='w')
    tk.Label(
        frm,
        text='Word report is rebuilt from scratch; stats and project updates come from an Excel workbook.',
        font=('Segoe UI', 10),
    ).pack(anchor='w', pady=(2, 4))
    if not HAVE_TKCALENDAR:
        tk.Label(
            frm,
            text='Tip: run "pip install tkcalendar" for a pop-up calendar picker (using plain text entry for now).',
            font=('Segoe UI', 9), fg='#8a6d00',
        ).pack(anchor='w', pady=(0, 10))
    else:
        tk.Frame(frm, height=10).pack()

    row = tk.Frame(frm); row.pack(fill='x', pady=5)
    tk.Label(row, text='Stats Excel', width=14, anchor='w').pack(side='left')
    tk.Entry(row, textvariable=excel_var, width=55).pack(side='left', padx=6)

    def browse():
        p = filedialog.askopenfilename(parent=root, title='Select weekly report Excel',
                                        filetypes=[('Excel workbook', '*.xlsx')])
        if p:
            excel_var.set(p)
    tk.Button(row, text='Browse...', command=browse).pack(side='left')

    grid = tk.Frame(frm); grid.pack(fill='x', pady=7)
    tk.Label(grid, text='Week Start', width=14, anchor='w').grid(row=0, column=0, sticky='w', pady=5)
    start_widget, get_start = _make_date_picker(grid, default_start)
    start_widget.grid(row=0, column=1, sticky='w')

    tk.Label(grid, text='Week End', width=14, anchor='w').grid(row=1, column=0, sticky='w', pady=5)
    end_widget, get_end = _make_date_picker(grid, default_start + timedelta(days=6))
    end_widget.grid(row=1, column=1, sticky='w')

    def sync_week_end(*_args):
        try:
            new_end = get_start() + timedelta(days=6)
            if HAVE_TKCALENDAR:
                end_widget.set_date(new_end)
            else:
                end_widget.delete(0, 'end')
                end_widget.insert(0, new_end.strftime('%d/%m/%Y'))
        except Exception:
            pass

    if HAVE_TKCALENDAR:
        start_widget.bind('<<DateEntrySelected>>', sync_week_end)
    tk.Button(grid, text='Set end = start + 6 days', command=sync_week_end).grid(row=0, column=2, padx=10)

    preview = tk.Text(frm, height=16, width=90, font=('Consolas', 9))
    preview.pack(fill='x', pady=15)
    preview.insert('1.0', 'No stats loaded.')
    preview.configure(state='disabled')

    loaded = {'stats': None, 'customer': [], 'internal': []}

    def show_preview(week_start, week_end, stats, customer, internal):
        lines = [f'Week {week_start.isocalendar()[1]}: {week_start:%d/%m/%Y} - {week_end:%d/%m/%Y}', '']
        lines.append(f"1. Smart Hands: {stats['smart_hands']} | {stats['smart_hands_hours']} hours")
        lines.append(f"2. Trouble Tickets: {stats['trouble_tickets']} | {stats['trouble_ticket_hours']} hours")
        lines.append(f"3. Cross connects installed: {stats['cross_connect_installed']}")
        lines.append(f"4. Cross connects de-installed: {stats['cross_connect_deinstalled']}")
        lines.append(f"5. Cross connects migrated: {stats['cross_connect_migrated']}")
        lines.append(f"6. Commit Compliance: {stats['commit_compliance']}%")
        lines.append(f"7. Labour: {stats['labour_hours']} hrs | {stats['labour_capitalization']}%")
        lines.append('')
        lines.append(f'Customer projects ({len(customer)}):')
        for p in customer:
            lines.append(f"  - {p['name']}" + (f" [{p['reference']}]" if p['reference'] else ''))
        lines.append('')
        lines.append(f'Internal projects ({len(internal)}):')
        for p in internal:
            lines.append(f"  - {p['name']}" + (f" [{p['reference']}]" if p['reference'] else ''))
        text = '\n'.join(lines)
        preview.configure(state='normal'); preview.delete('1.0', 'end'); preview.insert('1.0', text); preview.configure(state='disabled')

    def load_stats():
        if not excel_var.get():
            messagebox.showwarning('Excel required', 'Select the weekly report Excel workbook first.', parent=root)
            return None
        week_start = get_start()
        week_end = get_end()
        stats = read_stats_from_excel(excel_var.get(), week_start)
        customer = read_customer_projects(excel_var.get(), week_start)
        internal = read_internal_projects(excel_var.get(), week_start)
        loaded.update(stats=stats, customer=customer, internal=internal)
        show_preview(week_start, week_end, stats, customer, internal)
        return week_start, week_end, stats, customer, internal

    def load_stats_safe():
        try:
            return load_stats()
        except Exception as exc:
            messagebox.showerror('Error loading Excel', str(exc), parent=root)
            return None

    def generate():
        result = load_stats_safe()
        if result is None:
            return
        week_start, week_end, stats, customer, internal = result
        out = filedialog.asksaveasfilename(
            parent=root, title='Save weekly report', defaultextension='.docx',
            filetypes=[('Word document', '*.docx')],
            initialfile=f'SY1-2 - Weekly Report - Week {week_start.isocalendar()[1]} - {week_start:%Y-%m-%d}.docx',
        )
        if not out:
            return
        try:
            build_report(out, week_start, week_end, stats, customer, internal)
        except Exception as exc:
            messagebox.showerror('Error generating report', str(exc), parent=root)
            return
        messagebox.showinfo('Completed', f'Report generated:\n\n{out}', parent=root)

    buttons = tk.Frame(frm); buttons.pack(fill='x')
    tk.Button(buttons, text='Load / Preview Stats', command=load_stats_safe, padx=12, pady=7).pack(side='left')
    tk.Button(buttons, text='Generate Weekly Report', command=generate, font=('Segoe UI', 11, 'bold'), padx=14, pady=7).pack(side='right')

    tk.Label(
        frm,
        text='Excel structure: "Stats" sheet (one row per week) + "CustomerProjects" / "InternalProjects"\n'
             'sheets (one row per project item, tagged with the same Week Start date). See the template workbook.',
        fg='#555', justify='left',
    ).pack(anchor='w', pady=(14, 0))

    root.mainloop()


if __name__ == '__main__':
    gui()
