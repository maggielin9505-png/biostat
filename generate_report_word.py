import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    """
    kwargs: top, bottom, left, right
    values: dict(sz=12, val='single', color='FF0000')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        if border_name in kwargs:
            edge = OxmlElement(f'w:{border_name}')
            for key, val in kwargs[border_name].items():
                edge.set(qn(f'w:{key}'), str(val))
            tcBorders.append(edge)
    tcPr.append(tcBorders)

def add_heading_with_spacing(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    for r in h.runs:
        r.font.name = '微軟正黑體'
        r._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
        if level == 1:
            r.font.size = Pt(16)
            r.font.bold = True
            r.font.color.rgb = RGBColor(27, 54, 93) # #1B365D
        elif level == 2:
            r.font.size = Pt(13.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(41, 128, 185) # #2980B9
        elif level == 3:
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = RGBColor(52, 73, 94) # #34495E
    return h

def add_paragraph(doc, text="", style='Normal', space_after=5, line_spacing=1.15, bold=False, italic=False, color=None):
    p = doc.add_paragraph(text, style=style)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    for r in p.runs:
        r.font.name = '微軟正黑體'
        r._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
        r.font.size = Pt(10.5)
        if bold: r.font.bold = True
        if italic: r.font.italic = True
        if color: r.font.color.rgb = color
    return p

def add_callout_box(doc, text_list, title="重點提示"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border thick, others none
    set_cell_border(cell, 
                    left={'sz': 24, 'val': 'single', 'color': '1B365D'},
                    top={'sz': 4, 'val': 'single', 'color': 'D0D7DE'},
                    right={'sz': 4, 'val': 'single', 'color': 'D0D7DE'},
                    bottom={'sz': 4, 'val': 'single', 'color': 'D0D7DE'})
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"【{title}】\n")
    run_t.font.name = '微軟正黑體'
    run_t._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
    run_t.font.bold = True
    run_t.font.size = Pt(11)
    run_t.font.color.rgb = RGBColor(27, 54, 93)
    
    for item in text_list:
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(3)
        p2.paragraph_format.line_spacing = 1.15
        run = p2.add_run(item)
        run.font.name = '微軟正黑體'
        run._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(44, 62, 80)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def format_table(table, col_widths, col_alignments):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    for row_idx, row in enumerate(table.rows):
        is_header = (row_idx == 0)
        for col_idx, cell in enumerate(row.cells):
            cell.width = Inches(col_widths[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            
            # borders
            border_kwargs = {
                'top': {'sz': 4, 'val': 'single', 'color': 'CBD5E1'},
                'bottom': {'sz': 4, 'val': 'single', 'color': 'CBD5E1'},
                'left': {'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                'right': {'sz': 4, 'val': 'single', 'color': 'E2E8F0'}
            }
            if is_header:
                border_kwargs['top'] = {'sz': 12, 'val': 'single', 'color': '1B365D'}
                border_kwargs['bottom'] = {'sz': 12, 'val': 'single', 'color': '1B365D'}
            set_cell_border(cell, **border_kwargs)
            
            # background
            if is_header:
                set_cell_background(cell, "1B365D")
            else:
                if row_idx % 2 == 1:
                    set_cell_background(cell, "FFFFFF")
                else:
                    set_cell_background(cell, "F8FAFC")
            
            # paragraph
            for p in cell.paragraphs:
                p.alignment = col_alignments[col_idx]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    r.font.name = '微軟正黑體'
                    r._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
                    r.font.size = Pt(9.5) if not is_header else Pt(10)
                    if is_header:
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        r.font.color.rgb = RGBColor(30, 41, 59)

def build_word_document():
    doc = Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run("社區居民疾病流行與飲用水來源調查報告書")
    run_title.font.name = '微軟正黑體'
    run_title._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(27, 54, 93)
    
    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Epidemiological Investigation Report on Disease Outbreak and Water Sources")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(12)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 116, 139)
    
    # Meta box
    meta_p = doc.add_paragraph()
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_p.paragraph_format.space_after = Pt(16)
    run_meta = meta_p.add_run("調查對象：社區全體受訪居民 (N = 669)   |   分析工具：SPSS / Python Biostatistics   |   調查日期：2026年3月")
    run_meta.font.name = '微軟正黑體'
    run_meta._r.get_or_add_rPr().insert(0, parse_xml(r'<w:rFonts %s w:eastAsia="微軟正黑體"/>' % nsdecls('w')))
    run_meta.font.size = Pt(9.5)
    run_meta.font.color.rgb = RGBColor(71, 85, 105)
    
    # Horizontal rule
    p_hr = doc.add_paragraph()
    p_hr.paragraph_format.space_after = Pt(12)
    r_hr = p_hr.add_run("―" * 55)
    r_hr.font.color.rgb = RGBColor(203, 213, 225)
    p_hr.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 執行摘要
    add_heading_with_spacing(doc, "執行摘要 (Executive Summary)", level=1)
    
    summary_bullets = [
        "• 患病嚴重程度分佈：健康未發病者佔 23.02%（154人）、輕度患病（minor）佔 30.94%（207人）、重度患病（severe）佔 46.04%（308人），全社區總罹病率達 76.98%（515人）。患病者中重度患者佔近六成（59.81%）。",
        "• 飲用水來源種類比例：深井水（Deep well）佔 75.04%（502人）、淺井水（shallow well）佔 17.19%（115人）、自來水（Top water）佔 7.77%（52人）。水質二分群（water_group）中深井水組（Group 2）佔 75.04%，非深井水組（Group 1）佔 24.96%。",
        "• 疾病死亡率與致死特徵：全體粗死亡率為 76.98%（515人），存活率為 23.02%（154人）。關鍵發現：疾病個案致死率（Case Fatality Rate, CFR）高達 100.0%，所有 515 名染病者最終全數死亡，存活者皆為未患病之健康居民。",
        "• 人口學特徵（性別與年齡）：女性佔 61.14%（409人）、男性佔 38.86%（260人），女男比 1.57 : 1。平均年齡 39.67 ± 18.72 歲。卡方檢定證實性別（χ² = 0.001, p = 0.978）及年齡分組（χ² = 0.015, p = 0.903）與患病及死亡完全無關。",
        "• 致病因果推論與統計支持：飲用「深井水」為致病與致死的決定性危險因子。深井水組之患病死亡率高達 99.80%，相較非深井水組（8.38%），卡方檢定具極高度顯著性（χ² = 590.987, p < 0.001），風險比 RR = 11.905（95% CI: 7.211 ~ 19.655），勝算比 OR = 5,475.214（95% CI: 714.178 ~ 41,975.493）。自來水呈現 100% 絕對保護效果（0 人發病或死亡）。"
    ]
    add_callout_box(doc, summary_bullets, title="調查核心要點速覽")

    # 壹、調查背景與方法
    add_heading_with_spacing(doc, "壹、調查背景與研究方法", level=1)
    add_paragraph(doc, "本調查起源於某特定社區近期爆發大規模重大不明原因疾病，造成居民嚴重病變及高死亡率。為釐清疾病之流行病學特徵、確認可疑暴露源並評估風險強度，公共衛生調查小組對社區全體 669 名受訪居民進行普查，蒐集人口學背景（性別、年齡）、環境暴露史（主要日常飲用水來源）以及臨床病程結果（疾病嚴重程度、生存狀態）。")
    add_paragraph(doc, "統計分析採用 IBM SPSS Statistics 與 Python 統計模組進行交叉分析，推論統計包含 Pearson 卡方檢定（Chi-square test）、Yates 連續校正卡方檢定、概似比檢定與 Fisher 精確檢定，並計算相對風險比（Relative Risk, RR）與勝算比（Odds Ratio, OR）及其 95% 信賴區間。")

    # 貳、人口學特徵與描述性流行病學結果
    add_heading_with_spacing(doc, "貳、人口學特徵與描述性流行病學結果", level=1)
    
    # 1. 性別
    add_heading_with_spacing(doc, "一、性別比例分佈", level=2)
    add_paragraph(doc, "全社區 669 位居民中，性別分佈以女性居多，女性佔 61.14%，男性佔 38.86%，女男比例約為 1.57 : 1（每 100 位男性對應約 157 位女性）。")
    
    t_gender = doc.add_table(rows=3, cols=4)
    format_table(t_gender, [1.5, 1.5, 1.5, 2.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.LEFT])
    # Header
    for i, title in enumerate(["性別 (Gender)", "人數 (Count)", "佔比 (%)", "性別比例"]):
        t_gender.cell(0, i).paragraphs[0].text = title
    # Rows
    t_gender.cell(1, 0).paragraphs[0].text = "女性 (Female, F)"
    t_gender.cell(1, 1).paragraphs[0].text = "409"
    t_gender.cell(1, 2).paragraphs[0].text = "61.14%"
    t_gender.cell(1, 3).paragraphs[0].text = "女 : 男 = 1.57 : 1"
    
    t_gender.cell(2, 0).paragraphs[0].text = "男性 (Male, M)"
    t_gender.cell(2, 1).paragraphs[0].text = "260"
    t_gender.cell(2, 2).paragraphs[0].text = "38.86%"
    t_gender.cell(2, 3).paragraphs[0].text = "男 : 女 = 0.636 : 1"
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 2. 年齡
    add_heading_with_spacing(doc, "二、年齡分佈特徵", level=2)
    add_paragraph(doc, "全體居民年齡介於 8 至 90 歲之間，平均年齡為 39.67 ± 18.72 歲（95% 信賴區間：38.24 ~ 41.09 歲），中位數為 40.0 歲，四分位距（IQR）為 26 歲。經年齡分組後，≤40 歲組共有 349 人（52.17%），>40 歲組共有 320 人（47.83%），兩組人口比例相當接近。")

    # 3. 飲用水來源
    add_heading_with_spacing(doc, "三、飲用水來源種類比例", level=2)
    add_paragraph(doc, "調查之飲用水來源分為三種：深井水（Deep well）、淺井水（shallow well）與自來水（Top water）。同時依據水質特性劃分為 Group 2（深井水）與 Group 1（非深井水：淺井水+自來水）。")

    t_water = doc.add_table(rows=4, cols=5)
    format_table(t_water, [1.8, 1.2, 1.1, 1.2, 1.2], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    headers_water = ["飲用水種類", "水質組別", "人數", "佔全體比例 (%)", "累積比例 (%)"]
    for i, h in enumerate(headers_water):
        t_water.cell(0, i).paragraphs[0].text = h
    rows_water = [
        ["深井水 (Deep well)", "Group 2", "502", "75.04%", "75.0%"],
        ["淺井水 (shallow well)", "Group 1", "115", "17.19%", "92.2%"],
        ["自來水 (Top water)", "Group 1", "52", "7.77%", "100.0%"]
    ]
    for r_idx, r_data in enumerate(rows_water, start=1):
        for c_idx, val in enumerate(r_data):
            t_water.cell(r_idx, c_idx).paragraphs[0].text = val
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4. 患病嚴重程度
    add_heading_with_spacing(doc, "四、患病嚴重程度比例", level=2)
    add_paragraph(doc, "居民健康狀態分為健康未染病（Healthy）、輕度發病（minor）與重度發病（severe）。")

    t_dis = doc.add_table(rows=5, cols=4)
    format_table(t_dis, [1.8, 1.4, 1.6, 1.7], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    for i, h in enumerate(["患病嚴重程度", "人數 (Count)", "佔全體比例 (%)", "在患病者中佔比 (%)"]):
        t_dis.cell(0, i).paragraphs[0].text = h
    rows_dis = [
        ["健康 / 未發病 (Healthy)", "154", "23.02%", "―"],
        ["輕度患病 (minor)", "207", "30.94%", "40.19%"],
        ["重度患病 (severe)", "308", "46.04%", "59.81%"],
        ["患病者合計 (Diseased)", "515", "76.98%", "100.0%"]
    ]
    for r_idx, r_data in enumerate(rows_dis, start=1):
        for c_idx, val in enumerate(r_data):
            t_dis.cell(r_idx, c_idx).paragraphs[0].text = val
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 5. 疾病死亡率
    add_heading_with_spacing(doc, "五、疾病死亡率與致死特徵分析", level=2)
    add_paragraph(doc, "全社區居民之生存結果（survival outcome）中，死亡人數高達 515 人，佔全體居民的 76.98%；存活人數為 154 人，佔 23.02%。")
    add_paragraph(doc, "進一步交叉分析「患病嚴重程度」與「生存結果」，揭示了本疾病極端致命的流行特徵：")

    t_cross_surv = doc.add_table(rows=5, cols=5)
    format_table(t_cross_surv, [1.8, 1.2, 1.2, 1.1, 1.2], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    for i, h in enumerate(["患病嚴重程度", "存活人數 (alive)", "死亡人數 (Dead)", "組內合計", "組內死亡率"]):
        t_cross_surv.cell(0, i).paragraphs[0].text = h
    rows_cs = [
        ["健康未發病 (Healthy)", "154 (100.0%)", "0 (0.0%)", "154", "0.0%"],
        ["輕度患病 (minor)", "0 (0.0%)", "207 (100.0%)", "207", "100.0%"],
        ["重度患病 (severe)", "0 (0.0%)", "308 (100.0%)", "308", "100.0%"],
        ["全體總計", "154 (23.02%)", "515 (76.98%)", "669", "76.98%"]
    ]
    for r_idx, r_data in enumerate(rows_cs, start=1):
        for c_idx, val in enumerate(r_data):
            t_cross_surv.cell(r_idx, c_idx).paragraphs[0].text = val
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_paragraph(doc, "【流行病學關鍵致死特徵解析】", bold=True, color=RGBColor(27, 54, 93))
    add_paragraph(doc, "1. 粗死亡率 (Crude Mortality Rate)：全體 669 位居民中死亡 515 人，粗死亡率為 76.98%。")
    add_paragraph(doc, "2. 個案致死率 (Case Fatality Rate, CFR)：在 515 位患病個案中，死亡人數為 515 人，個案致死率達到 100.0%（515 / 515）。")
    add_paragraph(doc, "3. 凡是發病居民（不論輕度或重度），最終無一倖存全數死亡；全體存活者 154 人均為未曾罹病之健康者。因此在本研究中，「是否患病」與「是否死亡」在統計上呈現完全重合。")

    # 參、居民患病及死亡原因探討
    add_heading_with_spacing(doc, "參、居民患病及死亡原因探討與因果推論", level=1)
    add_paragraph(doc, "為探尋導致居民大量罹病及死亡的真正致病原因，研究團隊針對環境水質暴露進行多層次分析，並排除人口學潛在干擾因素：")
    
    add_heading_with_spacing(doc, "一、水質暴露之階梯式病因效應（劑量-反應梯度）", level=2)
    add_paragraph(doc, "交叉比對三種飲用水來源與患病嚴重程度，呈現極度顯著的劑量-效應暴露梯度（Dose-Response Gradient）：")

    t_water_dis = doc.add_table(rows=5, cols=6)
    format_table(t_water_dis, [1.6, 1.0, 1.0, 1.0, 1.0, 1.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    for i, h in enumerate(["飲用水來源", "健康人數", "輕度患病", "重度患病", "總人數", "罹病/死亡率"]):
        t_water_dis.cell(0, i).paragraphs[0].text = h
    rows_wd = [
        ["自來水 (Top water)", "52 (100.0%)", "0 (0.0%)", "0 (0.0%)", "52", "0.00%"],
        ["淺井水 (shallow well)", "101 (87.83%)", "14 (12.17%)", "0 (0.0%)", "115", "12.17%"],
        ["深井水 (Deep well)", "1 (0.20%)", "193 (38.45%)", "308 (61.35%)", "502", "99.80%"],
        ["全體總計", "154 (23.02%)", "207 (30.94%)", "308 (46.04%)", "669", "76.98%"]
    ]
    for r_idx, r_data in enumerate(rows_wd, start=1):
        for c_idx, val in enumerate(r_data):
            t_water_dis.cell(r_idx, c_idx).paragraphs[0].text = val
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_paragraph(doc, "【因果推論要點】")
    add_paragraph(doc, "1. 飲用自來水組（52人）：完全未發病、無人死亡，呈現 100% 絕對保護效果。")
    add_paragraph(doc, "2. 飲用淺井水組（115人）：罹病與死亡率為 12.17%（14人），且全數屬於輕度病況，無任何重度患者，存活率達 87.83%。")
    add_paragraph(doc, "3. 飲用深井水組（502人）：罹病與死亡率高達 99.80%（501人），且重度患者高達 308 人（佔深井水飲用者的 61.35%），僅 1 人倖免。")
    add_paragraph(doc, "4. 生物毒理與公共衛生對照：該流行特徵與臺灣歷史上西南沿海地區因飲用含高濃度無機砷之深層地下水而爆發之「烏腳病（Blackfoot Disease）」完全吻合。深井水抽自深層地質，極可能溶解了大量劇毒重金屬或特定有害物質，成為毒性最高、致死率最強的源頭。")

    add_heading_with_spacing(doc, "二、排除性別與年齡之干擾影響", level=2)
    add_paragraph(doc, "調查資料顯示，女性罹病死亡率為 77.02%（315/409），男性罹病死亡率為 76.92%（200/260），性別間死亡率差異僅 0.10%。在年齡方面，≤40 歲組死亡率為 76.79%（268/349），>40 歲組死亡率為 77.19%（247/320），組間差異僅 0.40%。單因子變異數分析亦證實年齡在各水質組及患病狀態組間皆無統計差異（p > 0.50）。證實疾病爆發純係環境飲用水暴露所致，非性別或年齡易感性造成。")

    # 肆、推論統計檢定資料
    add_heading_with_spacing(doc, "肆、推論統計檢定資料（卡方值、P值、風險比 RR、勝算比 OR）", level=1)
    add_paragraph(doc, "為提供完整嚴密的推論統計證據，研究建立列聯表並計算卡方統計量、P 值、風險比（RR）與勝算比（OR）：")

    add_heading_with_spacing(doc, "一、主要暴露檢定：深井水 (Group 2) vs 非深井水 (Group 1)", level=2)
    add_paragraph(doc, "以「非深井水組（自來水+淺井水）」為對照未暴露組，「深井水組」為暴露組，建立 2×2 列聯分析：")

    t_stat_main = doc.add_table(rows=7, cols=5)
    format_table(t_stat_main, [1.8, 1.2, 0.8, 1.2, 1.6], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER])
    for i, h in enumerate(["統計檢定項目", "統計數值", "自由度", "顯著性 P 值", "95% 信賴區間"]):
        t_stat_main.cell(0, i).paragraphs[0].text = h
    rows_sm = [
        ["Pearson 卡方檢定 (χ²)", "590.987", "1", "p < 0.001", "―"],
        ["連續校正卡方 (Yates')", "585.839", "1", "p < 0.001", "―"],
        ["概似比 (Likelihood Ratio)", "611.228", "1", "p < 0.001", "―"],
        ["費雪精確檢定 (Fisher's)", "―", "―", "p < 0.001", "―"],
        ["風險比 (Relative Risk, RR)", "11.905", "―", "―", "7.211 ～ 19.655"],
        ["勝算比 (Odds Ratio, OR)", "5,475.214", "―", "―", "714.178 ～ 41,975.493"]
    ]
    for r_idx, r_data in enumerate(rows_sm, start=1):
        for c_idx, val in enumerate(r_data):
            t_stat_main.cell(r_idx, c_idx).paragraphs[0].text = val
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_paragraph(doc, "【統計指標詳細公式與意涵】", bold=True, color=RGBColor(27, 54, 93))
    add_paragraph(doc, "• 暴露組（深井水）罹病死亡率：R1 = 501 / 502 = 99.8008%")
    add_paragraph(doc, "• 非暴露組（非深井水）罹病死亡率：R0 = 14 / 167 = 8.3832%")
    add_paragraph(doc, "• 相對風險比 (Relative Risk, RR)：RR = R1 / R0 = 0.998008 / 0.083832 = 11.9048 ≈ 11.905（95% CI: 7.211 ~ 19.655）。說明飲用深井水的居民，其罹病與死亡的風險是未飲用深井水居民的 11.91 倍。")
    add_paragraph(doc, "• 勝算比 (Odds Ratio, OR)：OR = (501 × 153) / (1 × 14) = 76653 / 14 = 5475.2143 ≈ 5475.214（95% CI: 714.178 ~ 41975.493）。說明飲用深井水居民之患病與死亡勝算是未飲用深井水居民的 5,475.21 倍。")
    add_paragraph(doc, "• SPSS 報表存活保護效應對應：SPSS 輸出之「Cohort survival outcome = alive」相對風險比為 RR_alive = (153/167) / (1/502) = 459.916（95% CI: 64.877 ~ 3260.359），證實非深井水水源對存活具有近 460 倍之保護效果。")

    add_heading_with_spacing(doc, "二、多類別卡方檢定與水質兩兩次群比較", level=2)
    add_paragraph(doc, "1. 飲用水三來源多群卡方檢定（3×2 交叉表）：Pearson χ² = 593.981，df = 2，p < 0.001（概似比 χ² = 622.244，p < 0.001）。證實三種水質來源之患病死亡率存在高度極顯著差異。")
    add_paragraph(doc, "2. 深井水 vs. 淺井水次群分析（2×2 表格）：Pearson χ² = 520.660（df = 1, p < 0.001），連續校正 χ² = 514.329（p < 0.001）。風險比 RR = (501/502) / (14/115) = 8.198（95% CI: 5.018 ~ 13.394），勝算比 OR = (501×101) / (1×14) = 3,614.357（95% CI: 469.958 ~ 27,797.314）。說明即使同為井水，深井水之罹病致死風險仍是淺井水的 8.20 倍。")
    add_paragraph(doc, "3. 淺井水 vs. 自來水次群分析（2×2 表格）：Pearson χ² = 6.910（df = 1, p = 0.0086），連續校正 χ² = 5.416（p = 0.0200），Fisher 精確檢定 p = 0.0056。淺井水罹病率（12.17%）仍顯著高於自來水（0%），自來水之水質純淨無毒。")

    add_heading_with_spacing(doc, "三、人口學對照檢定（陰性對照）", level=2)
    add_paragraph(doc, "針對性別與年齡進行對照卡方檢定，結果如下表：")

    t_neg = doc.add_table(rows=3, cols=6)
    format_table(t_neg, [1.5, 1.2, 0.8, 1.1, 1.0, 1.0], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    for i, h in enumerate(["對照檢定項目", "卡方值 (χ²)", "自由度", "P 值 (p-value)", "風險比 (RR)", "勝算比 (OR)"]):
        t_neg.cell(0, i).paragraphs[0].text = h
    rows_neg = [
        ["性別 (男 vs 女)", "0.001", "1", "0.978", "0.999", "0.995"],
        ["年齡分組 (>40 vs ≤40)", "0.015", "1", "0.903", "1.005", "1.023"]
    ]
    for r_idx, r_data in enumerate(rows_neg, start=1):
        for c_idx, val in enumerate(r_data):
            t_neg.cell(r_idx, c_idx).paragraphs[0].text = val
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_paragraph(doc, "由上表可見，性別與年齡的卡方值均接近 0，P 值均遠大於 0.05（分別為 0.978 與 0.903），RR 與 OR 均極度接近 1.0，完全排除為危險因子。")

    # 伍、結論與建議
    add_heading_with_spacing(doc, "伍、調查結論與公共衛生防治建議", level=1)
    
    add_heading_with_spacing(doc, "一、調查結論", level=2)
    add_paragraph(doc, "1. 疫情本質：本社區暴發之重大疾病為嚴重的環境水質介質中毒暴露事件，具有極高的個案致死率（100.0%）。")
    add_paragraph(doc, "2. 唯一病因：飲用「深井水」為致病與致死的決定性原因。深井水組居民患病與死亡風險是非深井水組的 11.91 倍，勝算比達 5,475.21 倍，呈現極強烈的因果效應強度與劑量-反應梯度。")
    add_paragraph(doc, "3. 防護效果：現代自來水系統之罹病與死亡率均為 0.0%，具完全防護性。")

    add_heading_with_spacing(doc, "二、具體公共衛生防治處置建議", level=2)
    recommendations = [
        "1. 立即全面封閉深水井設施：主管機關應緊急勒令封閉社區內所有深水井取水口，設置警示標誌並切斷抽水電力，全面嚴禁居民飲用或烹調使用深層地下水。",
        "2. 全面鋪設並普及自來水管網：基於自來水居民患病率 0% 之實證，公衛水務單位應緊急調度供水車供應日常飲用水，並以專案預算加速鋪設自來水公用管網，達成 100% 自來水供水覆蓋。",
        "3. 淺井水質管制與重金屬檢測：淺井水雖無重度病例，但仍有 12.17% 罹病率。應立即對全區淺井進行無機砷、重金屬（鉛、鎘、汞）及理化指標檢驗，未符標準者一律禁用。",
        "4. 受暴露高風險居民之長期健康監測與醫療介入：針對過去曾飲用深井水之居民建立專案健康追蹤檔案，進行尿液與毛髮重金屬殘留檢定，定期安排微血管、神經與內臟病變篩檢，早期發現潛在病變並給予臨床治療支援。"
    ]
    add_callout_box(doc, recommendations, title="公共衛生緊急處置與政策行動指引")

    # Save file
    output_filename = "社區居民疾病流行與飲用水來源調查報告書.docx"
    doc.save(output_filename)
    print(f"Document successfully created and saved as: {output_filename}")

if __name__ == '__main__':
    build_word_document()
