"""
Build professional business report with OLE-embedded Excel charts.
Charts are double-click editable in Excel.
"""
import win32com.client as win32
import datetime, os

output_dir = r'C:\Users\HP\OneDrive\Documents\Data-entry-projects\Task 3 - Business report word formatting'
output_path = os.path.join(output_dir, 'Professional_Business_Report_2024.docx')

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

doc = word.Documents.Add()

# ============================================================
# DIMENSIONS & HELPERS
# ============================================================
# 1 inch = 72 points
TPI = 72  # twips per inch? No, points per inch
CPI = 567  # twips per cm

sec = doc.Sections(1)
sec.PageSetup.TopMargin = 72
sec.PageSetup.BottomMargin = 72
sec.PageSetup.LeftMargin = 86.4
sec.PageSetup.RightMargin = 86.4

try:
    doc.Background.Fill.ForeColor.RGB = 0xF7FBDF  # FDFBF7 in BGR
except:
    pass

sel = word.Selection

# Chart type constants (from Excel XlChartType)
xlColumnClustered = 51
xlPie = 5
xlLine = 4
xlValue = 2

def rgb(r, g, b):
    """Convert RGB to BGR for Word COM"""
    return b * 65536 + g * 256 + r

NAVY    = rgb(0x1F, 0x3A, 0x5F)
GOLD    = rgb(0xC9, 0xA8, 0x4C)
WHITE   = rgb(0xFF, 0xFF, 0xFF)
DARK    = rgb(0x33, 0x33, 0x33)
GREY    = rgb(0x99, 0x99, 0x99)
LIGHT_BG = rgb(0xFD, 0xFB, 0xF7)
GOLD_BG  = rgb(0xFD, 0xF8, 0xEE)

def nl(n=1):
    for _ in range(n):
        sel.TypeParagraph()

def text(txt, bold=False, italic=False, size=11, color=DARK, align='left', sb=0, sa=6):
    if sb:
        sel.ParagraphFormat.SpaceBefore = sb
    sel.ParagraphFormat.SpaceAfter = sa
    sel.Font.Name = 'Calibri'
    sel.Font.Size = size
    sel.Font.Bold = bold
    sel.Font.Italic = italic
    sel.Font.Color = color
    sel.ParagraphFormat.Alignment = {
        'left': 0, 'center': 1, 'right': 2
    }.get(align, 0)
    sel.TypeText(txt)
    sel.Font.Bold = False
    sel.Font.Italic = False
    sel.Font.Color = DARK

def heading(txt, level=1):
    styles = {1: 'Heading 1', 2: 'Heading 2', 3: 'Heading 3'}
    sel.Style = doc.Styles(styles[level])
    sel.Font.Name = 'Calibri'
    sel.TypeText(txt)
    sel.TypeParagraph()
    sel.Style = doc.Styles('Normal')

def hr():
    sel.ParagraphFormat.SpaceBefore = 2
    sel.ParagraphFormat.SpaceAfter = 2
    sel.Font.Size = 6
    sel.Font.Color = GOLD
    sel.TypeText('_' * 80)
    sel.TypeParagraph()
    sel.Font.Size = 11
    sel.Font.Color = DARK

def page_break():
    sel.InsertBreak(7)  # wdPageBreak

def bullet(txt, size=11):
    sel.Range.ListFormat.ApplyBulletDefault()
    sel.Font.Name = 'Calibri'
    sel.Font.Size = size
    sel.Font.Color = DARK
    sel.TypeText(txt)
    sel.TypeParagraph()

def add_chart_ole(chart_type, categories, values, series_name='Series 1',
                  chart_title='Chart', width=450, height=280):
    """Insert an OLE-embedded Excel worksheet with a chart"""
    shape = sel.Range.InlineShapes.AddOLEObject(
        ClassType='Excel.Sheet',
        FileName=''
    )
    shape.Width = width
    shape.Height = height
    wb = shape.OLEFormat.Object
    ws = wb.Worksheets(1)
    ws.Name = 'Data'

    # Headers
    ws.Cells(1, 1).Value = 'Category'
    ws.Cells(1, 2).Value = series_name
    for i, cat in enumerate(categories, 2):
        ws.Cells(i, 1).Value = cat
    for i, val in enumerate(values, 2):
        ws.Cells(i, 2).Value = val

    # Create chart
    charts = wb.Charts()
    chart = charts.Add()
    chart.ChartType = chart_type
    chart.SetSourceData(ws.Range(f'A1:B{len(categories)+1}'))
    chart.HasTitle = True
    chart.ChartTitle.Text = chart_title
    chart.ChartTitle.Font.Name = 'Calibri'
    chart.ChartTitle.Font.Size = 12
    chart.ChartTitle.Font.Bold = True

    # Size chart area
    chart.ChartArea.Width = 400
    chart.ChartArea.Height = 250

    # Hide gridlines
    try:
        ax = chart.Axes(xlValue)
        ax.HasMajorGridlines = False
    except:
        pass

    # Move chart to be the active sheet view
    chart.Location(2, 'Chart1')  # xlLocationAsObject

    # Hide the worksheet, show chart
    ws.Visible = 2  # xlVeryHidden
    chart.Visible = True

    wb.Application.ActiveWindow.WindowState = -4140  # xlMinimized

    return shape

def add_ole_pie_chart(categories, values, title='Market Share'):
    """Insert an OLE pie chart"""
    shape = sel.Range.InlineShapes.AddOLEObject(
        ClassType='Excel.Sheet',
        FileName=''
    )
    shape.Width = 400
    shape.Height = 300
    wb = shape.OLEFormat.Object
    ws = wb.Worksheets(1)
    ws.Name = 'Data'

    ws.Cells(1, 1).Value = 'Region'
    ws.Cells(1, 2).Value = 'Share'
    for i, (cat, val) in enumerate(zip(categories, values), 2):
        ws.Cells(i, 1).Value = cat
        ws.Cells(i, 2).Value = val

    charts = wb.Charts()
    chart = charts.Add()
    chart.ChartType = xlPie
    chart.SetSourceData(ws.Range(f'A1:B{len(categories)+1}'))
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.ChartTitle.Font.Name = 'Calibri'
    chart.ChartTitle.Font.Size = 12
    chart.ChartTitle.Font.Bold = True
    chart.HasLegend = True
    chart.Legend.Font.Name = 'Calibri'
    chart.ApplyDataLabels()

    ws.Visible = 2
    return shape

# ============================================================
# COVER PAGE
# ============================================================
for _ in range(5): nl()

text('─' * 50, size=10, color=GOLD, align='center')
nl()
text('ANNUAL SALES', size=34, bold=True, color=NAVY, align='center')
nl()
text('PERFORMANCE REPORT', size=34, bold=True, color=NAVY, align='center')
nl()
text('2024', size=28, bold=True, color=GOLD, align='center')
nl()
text('Consolidated Regional Sales Analysis  •  Strategic Insights', size=13, color=GREY, align='center')
nl()
text('─' * 50, size=10, color=GOLD, align='center')
for _ in range(3): nl()
text(f'Prepared: {datetime.date.today().strftime("%B %d, %Y")}', size=10, color=GREY, align='center')
nl()
text('Classification: Confidential', size=10, color=GREY, align='center')
nl()
text('For Management Review Only', size=10, color=GREY, align='center')

# ============================================================
# TABLE OF CONTENTS
# ============================================================
page_break()
heading('Table of Contents', 1)
hr()

toc = [
    'Executive Summary', 'Performance Dashboard', 'Regional Revenue Analysis',
    'Market Share Overview', 'Growth Trends', 'North Region', 'South Region',
    'East Region', 'West Region', 'Central Region', 'Key Findings', 'Strategic Recommendations'
]
for i, t in enumerate(toc, 1):
    text(f'{i}.  {t}', size=12, color=NAVY, sa=3)

# ============================================================
# 1. EXECUTIVE SUMMARY
# ============================================================
page_break()
heading('1. Executive Summary', 1)
hr()

text('This report presents a comprehensive analysis of the company\'s sales performance across '
     'all major operating regions for the fiscal year 2024. The objective is to provide management '
     'with actionable insights into regional performance trends, growth opportunities, and areas '
     'requiring strategic attention.')
nl()
text('Sales increased in several regions while some regions experienced slower growth. Management '
     'requested a consolidated report to review trends and opportunities across the full portfolio. '
     'This document consolidates data from all seven regions, analyzing key performance indicators '
     'and identifying cross-regional patterns that inform the strategic recommendations outlined '
     'in the final section.')
nl()
text('The findings indicate strong overall growth momentum, with particular strength in the '
     'West and South regions. The North and Central regions present targeted improvement '
     'opportunities that, if addressed, could unlock significant additional revenue.',
     italic=True, color=GREY)

# ============================================================
# 2. PERFORMANCE DASHBOARD
# ============================================================
page_break()
heading('2. Performance Dashboard', 1)
hr()
text('Key performance indicators across all operating regions for Q4 2024.', italic=True, color=GREY)
nl()

# KPI 2x3 table
tbl = doc.Tables.Add(sel.Range, 2, 3)
tbl.Style = 'Table Grid'
kpi_items = [
    ('$27.0M', 'Consolidated Revenue', 'Q4 2024 Total', False),
    ('11.2%', 'Avg. YoY Growth', 'Company-Wide', False),
    ('26.4%', 'Highest Market Share', 'West Region', False),
    ('5.2%', 'Improvement Target', 'Central Region', True),
    ('4.6/5.0', 'Customer Satisfaction', 'South Region', False),
    ('45%', 'E-Commerce Growth', 'West Region YoY', True),
]
for idx, (val, label, sub, gold) in enumerate(kpi_items):
    r = idx // 3 + 1
    c = idx % 3 + 1
    cell = tbl.Cell(r, c)
    cell.Shading.BackgroundPatternColor = GOLD_BG
    clr = GOLD if gold else NAVY
    cell.Range.Paragraphs(1).Alignment = 1
    cell.Range.Font.Name = 'Calibri'
    cell.Range.Font.Size = 22
    cell.Range.Font.Bold = True
    cell.Range.Font.Color = clr
    cell.Range.Text = val
    cell.Range.InsertParagraphAfter()
    p2 = cell.Range.Paragraphs(2)
    p2.Alignment = 1
    p2.Range.Font.Size = 9
    p2.Range.Font.Color = DARK
    p2.Range.Text = label
    cell.Range.InsertParagraphAfter()
    p3 = cell.Range.Paragraphs(3)
    p3.Alignment = 1
    p3.Range.Font.Size = 8
    p3.Range.Font.Color = GOLD
    p3.Range.Text = sub

nl()

# Progress bars
text('Regional Goal Achievement', size=12, bold=True, color=NAVY)
nl()
text('Progress toward FY 2024 revenue targets by region:', italic=True, color=GREY, size=9)
nl()

progress = [('West', 92), ('South', 88), ('East', 79), ('North', 71), ('Central', 58)]
tbl2 = doc.Tables.Add(sel.Range, 6, 3)
tbl2.Style = 'Table Grid'
for j, h in enumerate(['Region', 'Progress', 'Achievement']):
    c = tbl2.Cell(1, j+1)
    c.Range.Font.Name = 'Calibri'
    c.Range.Font.Size = 9
    c.Range.Font.Bold = True
    c.Range.Font.Color = WHITE
    c.Range.ParagraphFormat.Alignment = 1
    c.Shading.BackgroundPatternColor = NAVY
    c.Range.Text = h

for i, (reg, pct) in enumerate(progress):
    r = i + 2
    c1 = tbl2.Cell(r, 1)
    c1.Range.Font.Name = 'Calibri'
    c1.Range.Font.Size = 10
    c1.Range.Font.Bold = True
    c1.Range.Font.Color = NAVY
    c1.Range.Text = reg
    c1.Width = 86.4

    filled = pct // 5
    empty = 20 - filled
    bar = '█' * filled + '░' * empty
    c2 = tbl2.Cell(r, 2)
    c2.Range.Font.Name = 'Calibri'
    c2.Range.Font.Size = 10
    c2.Range.Font.Color = NAVY
    c2.Range.Text = f'  {bar}'
    c2.Width = 259.2

    c3 = tbl2.Cell(r, 3)
    c3.Range.Font.Name = 'Calibri'
    c3.Range.Font.Size = 11
    c3.Range.Font.Bold = True
    c3.Range.Font.Color = NAVY
    c3.Range.ParagraphFormat.Alignment = 1
    c3.Range.Text = f'{pct}%'
    c3.Width = 72

    bg = GOLD_BG if i % 2 == 0 else WHITE
    for ci in range(3):
        tbl2.Cell(r, ci+1).Shading.BackgroundPatternColor = bg

nl()

# ============================================================
# 3. REGIONAL REVENUE ANALYSIS (with bar chart)
# ============================================================
page_break()
heading('3. Regional Revenue Analysis', 1)
hr()

text('The bar chart below illustrates Q4 2024 revenue performance across all five operating regions. '
     'The West region leads with $7.3M in revenue, followed closely by the South at $6.8M. Together, '
     'these two regions account for over half of total company revenue.')
nl()
text('Chart 1: Q4 2024 Revenue by Region (USD Millions)', bold=True, size=9, color=NAVY, align='center')
nl()

# Embedded bar chart
regions = ['North', 'South', 'East', 'West', 'Central']
revenues = [4.2, 6.8, 5.1, 7.3, 3.6]
add_chart_ole(xlColumnClustered, regions, revenues, 'Revenue ($M)', 'Q4 2024 Revenue by Region')
nl()

# Revenue table
text('Revenue Breakdown by Region', size=12, bold=True, color=NAVY)
nl()

tbl3 = doc.Tables.Add(sel.Range, 7, 5)
tbl3.Style = 'Table Grid'
headers = ['Region', 'Q4 Revenue', 'YoY Growth', 'Market Share', 'Status']
rev_data = [
    ['West', '$7.3M', '+16.8%', '26.4%', 'Exceeding'],
    ['South', '$6.8M', '+14.2%', '24.6%', 'Exceeding'],
    ['East', '$5.1M', '+11.3%', '20.1%', 'On Track'],
    ['North', '$4.2M', '+8.5%', '18.2%', 'On Track'],
    ['Central', '$3.6M', '+5.2%', '10.7%', 'Attention'],
    ['Total', '$27.0M', '+11.2%', '100%', '●'],
]
status_colors = {'Exceeding': rgb(0xD5, 0xF5, 0xE3), 'On Track': rgb(0xFD, 0xEB, 0xD0), 'Attention': rgb(0xFA, 0xDB, 0xD8)}

for j, h in enumerate(headers):
    c = tbl3.Cell(1, j+1)
    c.Range.Font.Name = 'Calibri'
    c.Range.Font.Size = 9
    c.Range.Font.Bold = True
    c.Range.Font.Color = WHITE
    c.Range.ParagraphFormat.Alignment = 1
    c.Shading.BackgroundPatternColor = NAVY
    c.Range.Text = h

for i, row in enumerate(rev_data):
    for j, val in enumerate(row):
        c = tbl3.Cell(i+2, j+1)
        c.Range.Font.Name = 'Calibri'
        c.Range.Font.Size = 10
        c.Range.Font.Color = NAVY
        c.Range.ParagraphFormat.Alignment = 1
        is_tot = (i == 5)
        if is_tot:
            c.Range.Font.Bold = True
            c.Range.Font.Color = GOLD
            c.Shading.BackgroundPatternColor = NAVY
        c.Range.Text = val
        if j == 4 and not is_tot:
            c.Shading.BackgroundPatternColor = status_colors.get(val, WHITE)
        elif not is_tot and i % 2 == 0 and j != 4:
            c.Shading.BackgroundPatternColor = GOLD_BG

nl()

# ============================================================
# 4. MARKET SHARE (pie chart)
# ============================================================
page_break()
heading('4. Market Share Overview', 1)
hr()

text('The West and South regions together command 51% of total company revenue. The pie chart below '
     'illustrates the market share distribution across all operating regions.')
nl()
text('Chart 2: Market Share Distribution by Region', bold=True, size=9, color=NAVY, align='center')
nl()

# Embedded pie chart
add_ole_pie_chart(regions, [18.2, 24.6, 20.1, 26.4, 10.7], 'Market Share by Region')
nl()

text('The West region holds the largest share at 26.4%, followed by the South at 24.6%. '
     'The Central region, while smallest at 10.7%, represents a significant growth opportunity.',
     italic=True, color=GREY)

# ============================================================
# 5. GROWTH TRENDS (bar chart with avg line)
# ============================================================
page_break()
heading('5. Growth Trends', 1)
hr()

text('Year-over-year growth rates vary significantly across regions. The company-wide average of '
     '11.2% was exceeded by the West (+16.8%) and South (+14.2%) regions.')
nl()
text('Chart 3: Year-over-Year Growth Rate by Region (%)', bold=True, size=9, color=NAVY, align='center')
nl()

# Add growth chart as dual-series bar chart
growth_vals = [8.5, 14.2, 11.3, 16.8, 5.2]

shape3 = sel.Range.InlineShapes.AddOLEObject(ClassType='Excel.Sheet', FileName='')
shape3.Width = 450
shape3.Height = 280
wb3 = shape3.OLEFormat.Object
ws3 = wb3.Worksheets(1)
ws3.Name = 'Data'
ws3.Cells(1, 1).Value = 'Region'
ws3.Cells(1, 2).Value = 'Growth (%)'
ws3.Cells(1, 3).Value = 'Company Avg'
for i, reg in enumerate(regions, 2):
    ws3.Cells(i, 1).Value = reg
    ws3.Cells(i, 2).Value = growth_vals[i-2]
    ws3.Cells(i, 3).Value = 11.2

c3 = wb3.Charts().Add()
c3.ChartType = xlColumnClustered
c3.SetSourceData(ws3.Range('A1:C6'))
c3.HasTitle = True
c3.ChartTitle.Text = 'YoY Growth Rate by Region'
c3.ChartTitle.Font.Name = 'Calibri'
c3.ChartTitle.Font.Size = 12
c3.ChartTitle.Font.Bold = True
# Change second series to line
c3.SeriesCollection(2).ChartType = xlLine
c3.SeriesCollection(2).Format.Line.ForeColor = rgb(0xC9, 0xA8, 0x4C)
c3.HasLegend = True
c3.Legend.Font.Name = 'Calibri'

ws3.Visible = 2
nl()

text('Key observation: Three regions (West, South, East) outperformed or matched the company average, '
     'while North and Central regions fell below the benchmark. Targeted interventions for these '
     'regions could unlock an estimated $2-3M in additional annual revenue.',
     italic=True, color=GREY)

# ============================================================
# 6-10. REGIONAL SECTIONS
# ============================================================
region_data = [
    ('6', 'North Region', 'North', 4.2, 8.5, 18.2, 71,
     'The North region demonstrated stable growth throughout FY 2024, with quarterly revenue '
     'increasing steadily across all four quarters. Sales increased in several sub-regions while '
     'some areas experienced relatively slower growth.',
     ['Q4 revenue of $4.2M (18.2% of total)',
      'YoY growth of 8.5%, below company average of 11.2%',
      'Customer acquisition cost reduced by 6%',
      'Two new distribution partnerships in Q3']),
    ('7', 'South Region', 'South', 6.8, 14.2, 24.6, 88,
     'The South region was among the strongest performers this year, posting double-digit growth '
     'and expanding market share. Sales increased significantly across most product lines.',
     ['Q4 revenue of $6.8M (24.6% of total)',
      'YoY growth of 14.2%, well above company average',
      'Customer satisfaction improved to 4.6/5.0',
      'Three major enterprise accounts closed in Q4']),
    ('8', 'East Region', 'East', 5.1, 11.3, 20.1, 79,
     'The East region maintained steady growth throughout the year, with consistent quarter-over-quarter '
     'improvements. Sales increased in urban markets while rural expansion gained traction.',
     ['Q4 revenue of $5.1M (20.1% of total)',
      'YoY growth of 11.3%, aligned with company average',
      'Rural market penetration increased 22% YoY',
      'Sales team productivity improved 15%']),
    ('9', 'West Region', 'West', 7.3, 16.8, 26.4, 92,
     'The West region delivered exceptional performance in FY 2024, achieving the highest growth rate '
     'across all regions. Sales increased substantially across both existing and new product categories.',
     ['Q4 revenue of $7.3M (26.4% of total)',
      'YoY growth of 16.8% — fastest growing region',
      'New product lines contributed 28% of revenue',
      'E-commerce channel grew 45%']),
    ('10', 'Central Region', 'Central', 3.6, 5.2, 10.7, 58,
     'The Central region showed moderate growth with room for improvement. Sales increased at a '
     'slower pace compared to other regions, reflecting both market challenges and execution gaps.',
     ['Q4 revenue of $3.6M (10.7% of total)',
      'YoY growth of 5.2% — below company average',
      'Cost optimization identified $420K in savings',
      'New regional director appointed in Q4']),
]

for num, title, short, rev, gr, ms, prog, body, highlights in region_data:
    page_break()
    heading(f'{num}. {title} Analysis', 1)
    hr()

    text(body)
    nl()
    text('Key Highlights', size=12, bold=True, color=NAVY)
    nl()
    for h in highlights:
        bullet(h)

    # Progress bar
    nl()
    text('Goal Achievement', size=11, bold=True, color=NAVY)
    nl()
    filled = prog // 5
    empty = 20 - filled
    bar = '█' * filled + '░' * empty
    tbl_p = doc.Tables.Add(sel.Range, 1, 2)
    tbl_p.Style = 'Table Grid'
    c1 = tbl_p.Cell(1, 1)
    c1.Range.Font.Name = 'Calibri'
    c1.Range.Font.Size = 10
    c1.Range.Font.Bold = True
    c1.Range.Font.Color = NAVY
    c1.Range.Text = f'{short} Region'
    c1.Width = 108
    c1.Shading.BackgroundPatternColor = GOLD_BG
    c2 = tbl_p.Cell(1, 2)
    c2.Range.Font.Name = 'Calibri'
    c2.Range.Font.Size = 10
    c2.Range.Font.Color = NAVY
    c2.Range.Text = f'  {bar}  {prog}%'
    c2.Width = 324
    c2.Shading.BackgroundPatternColor = GOLD_BG
    nl()

# ============================================================
# 11. KEY FINDINGS
# ============================================================
page_break()
heading('11. Key Findings', 1)
hr()

findings = [
    ('Strong Overall Growth',
     'The company achieved 11.2% year-over-year revenue growth in FY 2024, driven by strong '
     'performance in the West (+16.8%) and South (+14.2%) regions. Total consolidated revenue '
     'reached $27.0M for Q4 alone.'),
    ('Regional Performance Gap',
     'A notable gap exists between high-growth regions (West, South) and lower-growth regions '
     '(Central at +5.2%). Addressing this could unlock an estimated $2-3M in potential upside.'),
    ('Digital Channel Acceleration',
     'E-commerce channels are becoming increasingly important, with the West region\'s online '
     'channel growing 45% year-over-year. Digital transformation should be accelerated region-wide.'),
    ('Market Position Strength',
     'The company maintains strong market positions across all regions, with share from 10.7% '
     '(Central) to 26.4% (West). Expansion opportunities exist in North and Central regions.'),
]
for t, b in findings:
    heading(t, 2)
    text(b)

# ============================================================
# 12. STRATEGIC RECOMMENDATIONS
# ============================================================
page_break()
heading('12. Strategic Recommendations', 1)
hr()

text('Based on the analysis of regional performance data and emerging market trends, management '
     'recommends the following strategic actions for FY 2025:', italic=True)
nl()

recs = [
    ('Accelerate Digital Transformation',
     'Invest in e-commerce and digital marketing across all regions, focusing on North and Central '
     'regions where digital penetration lags. Target: 30% digital channel contribution by Q4 FY 2025.'),
    ('Regional Turnaround — Central',
     'Implement a dedicated growth initiative including leadership changes, targeted marketing, '
     'and distribution expansion. Target: 10%+ growth rate by Q2 FY 2025.'),
    ('Scale High-Performance Practices',
     'Document and replicate successful strategies from West and South regions across the '
     'organization — sales training, retention frameworks, and campaign templates.'),
    ('Cross-Regional Collaboration',
     'Establish a Regional Excellence Council for knowledge sharing between regions. '
     'Monthly cross-regional reviews to track progress against targets.'),
    ('Enhance Analytics Capabilities',
     'Deploy advanced analytics tools for real-time sales dashboards across all regions, '
     'enabling predictive forecasting and early warning systems.'),
]
for i, (t, b) in enumerate(recs, 1):
    heading(f'{i}. {t}', 2)
    text(b)

# ============================================================
# FOOTER
# ============================================================
for s in doc.Sections:
    ft = s.Footers(1)  # wdHeaderFooterPrimary
    ft.Range.Font.Name = 'Calibri'
    ft.Range.Font.Size = 8
    ft.Range.Font.Color = GREY
    ft.Range.ParagraphFormat.Alignment = 1
    ft.Range.Text = 'Page '
    # Add page number field
    ft.Range.Collapse(0)  # wdCollapseEnd
    word.Selection.Fields.Add(ft.Range, 17)  # wdFieldPage
    ft.Range.Collapse(0)
    ft.Range.Text = ' of '
    ft.Range.Collapse(0)
    word.Selection.Fields.Add(ft.Range, 42)  # wdFieldNumPages

# ============================================================
# SAVE
# ============================================================
try:
    os.remove(output_path)
except:
    pass
doc.SaveAs(output_path)
doc.Close()
word.Quit()

print(f'Done! Saved to:\n{output_path}')
print('Charts are embedded as editable Excel OLE objects — double-click to edit.')
