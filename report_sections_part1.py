# -*- coding: utf-8 -*-
"""
Part 1 of PISA 2025 School-based Curriculum Reform Report:
- Cover / Executive Header
- Executive Summary
- Section 1: Data Retrieval & Quantitative Methodology Guidelines
- Section 2: Curriculum Intelligence Framework (8 Dimensions)
"""

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from build_report import (
    HEX_NAVY, HEX_SLATE, HEX_DARK, HEX_MUTED, HEX_LIGHT_BG, HEX_ROW_ALT,
    HEX_ACCENT_RED, HEX_ACCENT_GOLD, HEX_WHITE, RGB_NAVY, RGB_SLATE, RGB_DARK,
    add_styled_heading, add_para, add_bullet, add_callout, add_styled_table,
    set_cell_background, set_cell_margins
)

def build_cover_and_header(doc):
    # Top Institutional Tag
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(0)
    p_top.paragraph_format.space_after = Pt(4)
    r_top = p_top.add_run("學校課程發展委員會 · 內部行政與決策專題研討報告")
    r_top.font.name = "微軟正黑體"
    r_top.font.size = Pt(9.5)
    r_top.font.color.rgb = RGBColor(113, 128, 150)
    r_top.bold = True

    # Main Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("從國際標竿到校本課堂：PISA 2025 澳門數據對學校課程改革之深度診斷與行動藍圖")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(18)
    r_title.bold = True
    r_title.font.color.rgb = RGB_NAVY

    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("在高成就與高公平背後：剖析深層素養斷層、學習動機危機與 AI 時代課程重構")
    r_sub.font.name = "微軟正黑體"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGB_SLATE

    # Meta Info Card (Table)
    tbl_meta = doc.add_table(rows=2, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_widths = [3.2, 3.3]
    for row in tbl_meta.rows:
        for i, w in enumerate(meta_widths):
            row.cells[i].width = Inches(w)
            set_cell_background(row.cells[i], HEX_LIGHT_BG)
            set_cell_margins(row.cells[i], top=70, bottom=70, left=100, right=100)

    p0 = tbl_meta.cell(0, 0).paragraphs[0]
    p0.add_run("呈交單位：").bold = True
    p0.add_run("課程發展處 / 數學科暨跨學科課程發展組")
    
    p1 = tbl_meta.cell(0, 1).paragraphs[0]
    p1.add_run("呈交對象：").bold = True
    p1.add_run("校長室、教務委員會、學術委員會、各學科組長")

    p2 = tbl_meta.cell(1, 0).paragraphs[0]
    p2.add_run("報告視角：").bold = True
    p2.add_run("課程發展助理主任（Curriculum Development Assistant Director）")

    p3 = tbl_meta.cell(1, 1).paragraphs[0]
    p3.add_run("報告日期：").bold = True
    p3.add_run("2026 年 9 月 | 文件層級：校內重要政策研討")

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(8)
    p_sp.paragraph_format.space_after = Pt(8)

def build_executive_summary(doc):
    add_styled_heading(doc, "【執行摘要】由「總分排名思維」全面轉向「課程智慧（Curriculum Intelligence）」", level=1)
    
    add_para(doc, "經濟合作與發展組織（OECD）於 2026 年 9 月 8 日正式公布 PISA 2025 國際評估結果。澳門 15 歲學生在本次評估中展現出令人矚目的總體成就：數學得分 549 分（全球第 3）、科學得分 541 分（全球第 3）、閱讀得分 501 分（全球第 5）；尤為亮眼的是，本屆首次作為核心創新領域呈現的「在數碼世界中的學習」（Learning in the Digital World: Computational Problem Solving, CPS）澳門勇奪 572 分，位列全球所有參測經濟體第一名，高表現學生（Level 5–6）佔比高達 53.7%。然而，作為學校課程發展管理者，若僅停留在「為澳門排名喝彩」的宏觀敘事，將嚴重錯失推動校本深層變革的關鍵契機。")
    
    add_para(doc, "深入穿透平均分表象，本報告基於 OECD 原始數據與澳門本地實證，挖掘出五大不容忽視的結構性隱患與悖論：")
    add_bullet(doc, "澳門 2022 至 2025 年數理表現維持平穩（數學 -3 分，科學 -2 分），但閱讀素養出現顯著滑落（-9 分，由 510 降至 501）。此滑落直接暴露出學生在面對跨學科長題幹、多模態圖表與複雜真實情境時的「學科素養（Disciplinary Literacy）」存在明顯短板，直接威脅高階數理建模之根基。", bold_prefix="1. 數理高光下的閱讀斷層：")
    add_bullet(doc, "雖然澳門社經背景對成績的解釋率僅約 4%（OECD 平均 12%），具備卓越的教育公平指標，但 2015 至 2025 年間優勢與弱勢學生的科學差距已現擴大趨勢；2022 至 2025 年數學最高 10%（P90）與最低 10%（P10）之分化顯著加劇，出現「高均分掩蓋下的底線下墜與兩極分化」隱憂。", bold_prefix="2. 頂層與底層差距擴大：")
    add_bullet(doc, "澳門僅 58.1% 的學生表示「喜歡在學校學習新事物」（OECD 平均 68.8%），學校歸屬感與家庭支持感指數相對低落，呈現典型的「很會應試、但缺乏求知熱忱」的情感沙漠現象，自主學習動機與成長型思維亟需制度化重塑。", bold_prefix="3. 高成就 × 低投入文化：")
    add_bullet(doc, "澳門 67% 的學生每週使用 AI Chatbot 輔助學習（OECD 平均 46%），但在作業初稿（40%）與文本摘要（41%）的高頻使用中，極易引發「認知外包（Cognitive Offloading）」與思維萎縮風險；同時，課堂數位干擾雖僅佔 13%，但受干擾學生在控制社經背景後科學成績驟降 24 分（OECD 僅 11 分），干擾代價極其高昂。", bold_prefix="4. AI 普及與數位干擾代價：")
    add_bullet(doc, "若將 15 歲需要具備的批判推理、系統建模與運算思維留至初三才突擊，為時已晚。學校必須啟動 P1–S6 逆向設計，打通小初高學段壁壘，嚴防各關鍵年段的能力斷層。", bold_prefix="5. 縱向課程貫通迫在眉睫：")

    add_callout(doc, [
        "本報告旨在為校長室、教務會與學科組提供一套具備操作性的「PISA Curriculum Intelligence Framework」，從外部標竿對比、認知能力反推、校本評量審計到 P1–S6 縱向重構，提供 10 大關鍵議題矩陣與行動路徑。",
        "核心策略倡議：校本評量告別單一均分，全面追蹤分佈與 P10/P90 離散度；將閱讀視為所有學科之基礎素養；在 AI 時代以 Q-V-R-E-R 思維五步法引導深度學習；構建 MTSS 多層級支援架構，落實「提升上限、守住底線」。"
    ], title="課程發展助理主任戰略倡議", alert=False)

def build_section_1(doc):
    add_styled_heading(doc, "第一章：PISA 數據源體系檢索指南與量化研究方法學規範", level=1)
    
    add_styled_heading(doc, "1.1 數據檢索的三個層級與政策語境對照", level=2)
    add_para(doc, "要落實真正證據導向（Evidence-based）的校本課程改革，科組長與管理團隊必須精準掌握 OECD 原始數據的架構層級，避免道聽途說或誤用未經加權的片面數據。研究體系應嚴格劃分為三個層級：")

    headers_l = ["資料層級", "核心資料庫 / 文件入口", "適用分析用途", "管理層優先度", "技術難度"]
    data_l = [
        ["Level 1 政策摘要", "OECD Macau Country Note\nOECD Education GPS - Macao", "宏觀掌握澳門歷年趨勢、強弱項領域、公平指標、AI 與校園環境摘要", "★★★★★", "★☆☆☆☆\n（無需代碼）"],
        ["Level 2 跨國統計", "OECD StatLink Excel Tables\nPISA 2025 Volume I-III 附表", "進行東亞標竿經濟體（星、日、韓、台、港）之均分、分位數與標準差精確比較", "★★★★★", "★★☆☆☆\n（表格檢索）"],
        ["Level 3 微觀數據", "PISA 2025 Database (PUFs)\nStudent/School/Teacher PUF", "校本研究人員自行建立多層次模型（HLM）、調節效應分析、認知作答時間對比", "★★★★☆", "★★★★★\n（需統計編程）"]
    ]
    add_styled_table(doc, headers_l, data_l, col_widths=[1.3, 1.8, 2.0, 0.7, 0.7])

    add_para(doc, "在引用 Level 1 與 Level 2 資料時，必須緊扣澳門教育及青年發展局（DSEDJ）於 2026 年 9 月 8 日發布的官方政策解讀。教青局明確指出：澳門 PISA 2025 成果得益於長年推動的跨學科課程改革、智慧校園建設以及人工智能與編程教育普及。這為本校將 PISA 指標轉化為正規課程政策提供了直接的體制背書與政策順風。")

    add_styled_heading(doc, "1.2 學校進行 PISA 微觀數據（PUFs）二次分析的方法學五大戒律", level=2)
    add_para(doc, "若本校研究團隊或科組長計劃下載 PISA Public Use Files（PUFs）針對特定題目、學生性別、家庭社經地位進行深度交叉分析，必須嚴格遵守 OECD 大規模國際評估（Large-Scale Assessment）的計量統計規範，切忌將 PUF 數據當作一般常態分佈樣本直接以 Excel 計算平均值或跑簡單線性迴歸。違背以下五大戒律將導致嚴重的估計偏誤與虛假顯著性：")

    add_bullet(doc, "PISA 採用兩階段分層抽樣（先抽學校，再在校內隨機抽取 15 歲學生），非簡單隨機抽樣。學生並非獨立同分佈，校內同質性顯著，必須納入集群效應（Clustering Effect）。", bold_prefix="① 嚴禁忽視複雜抽樣設計：")
    add_bullet(doc, "每位學生代表母體的機率不同。所有描述性統計與模型估算均須加權計算，必須強制使用最終學生權重變量（Final Student Weight: W_FSTUWT），否則將嚴重扭曲全澳真實母體估計值。", bold_prefix="② 嚴禁忽略學生權重（Student Weights）：")
    add_bullet(doc, "計算標準誤（Standard Errors）時不可採用常態公式，必須使用 OECD 提供的 80 個重複抽樣權重（Replicate Weights: W_FSTR1 至 W_FSTR80），並採用 Fay 調整之平衡重複抽樣法（Fay's Balanced Repeated Replication, BRR）以確保推論統計的可靠性。", bold_prefix="③ 嚴格執行平衡重複抽樣（Fay's BRR）：")
    add_bullet(doc, "PISA 並非對每位學生進行全卷施測，而是採用矩陣抽樣（Matrix Sampling）結合項目反應理論（IRT），每位學生僅回答部分題目。學生成績並非單一確定分數，而是由 10 組合理值（Plausible Values: PV1 至 PV10）構成。嚴禁將 10 個 PV 直接取平均後當作因變量跑 OLS 迴歸！標準做法是分別對 10 組 PV 建立 10 次統計模型，再依照 Rubin's Rules 將參數與變異數合併計算。", bold_prefix="④ 嚴禁對單一合理值（PV）進行推論：")
    add_bullet(doc, "建議校內數據分析團隊統一採用專門處理大規模評估的套件，例如 R 語言之 intsvy 或 edsurvey 套件、SAS 官方巨集或 SPSS 專用模組。需注意 PISA 2025 Technical Report 目前處於 Draft 階段，最終官方技術手冊預計於 2027 年定稿發布，研究中涉及的新指標需持續核對校正。", bold_prefix="⑤ 規範計算工具與技術版本核驗：")

def build_section_2(doc):
    add_styled_heading(doc, "第二章：學校 PISA 課程智慧架構（Curriculum Intelligence Framework）8 大維度", level=1)
    add_para(doc, "為了讓全體科組長與教師擺脫「只看排行榜」的狹隘視野，課程發展處特構建「學校 PISA 課程智慧架構（Curriculum Intelligence Framework）」。該架構涵蓋 8 個核心維度，將宏觀評估指標精準對接至微觀課堂與課程機制：")

    headers_dim = ["分析維度", "核心追問", "PISA 2025 澳門關鍵數據", "OECD 參照基準", "對本校課程改革之具體作用"]
    data_dim = [
        [
            "① 學業成就\n(Achievement)",
            "學生究竟學到了多少？各核心領域成效如何？",
            "數學：549 分 (全球第3)\n科學：541 分 (全球第3)\n閱讀：501 分 (全球第5)\n運算解難(CPS)：572 (第1)",
            "數學：463 分\n科學：482 分\n閱讀：461 分\nCPS：500 分",
            "驗證現行數理與資訊科技課程效能，確保持續領先優勢，重新定位各領域資源配置比重。"
        ],
        [
            "② 分佈結構\n(Distribution)",
            "哪些學生在掉隊？高低分層分化程度如何？",
            "科學低成就(<L2)：7.9%\n科學頂尖(L5-6)：14.5%\n數學頂尖(L5-6)：約 29%\nCPS 頂尖(L5-6)：53.7%",
            "科學 <L2：24%\n科學 L5-6：7%\n數學 L5-6：9%\nCPS L5-6：約 15%",
            "打破全校均分盲區，建立校內「P10/P90 離散度」動態追蹤機制，落實分層補底與高階拔尖。"
        ],
        [
            "③ 核心素養\n(Competency)",
            "學生是機械記憶還是具備真實情境解難能力？",
            "數學建模、科學論證推理、系統分解與除錯表現強勁；長文本非連續閱讀稍弱",
            "OECD 平均在情境建模與運算探索題型上表現普遍受限",
            "重構學科命題範式：由「技巧公式刷題」全面轉向「情境解讀 → 數學建模 → 運算 → 評估反思」。"
        ],
        [
            "④ 歷年趨勢\n(Trend)",
            "學生表現是持續精進、高原停滯還是出現警訊？",
            "科學：較22年 -2分 (持平)\n數學：較22年 -3分 (持平)\n閱讀：較22年 -9分 (顯著下滑)",
            "OECD 歷年平均呈緩慢下降趨勢，疫後復甦普遍乏力",
            "確立警示性重大課題：閱讀能力滑落非語文科單科之責，需全面啟動跨學科「學科素養閱讀」改革。"
        ],
        [
            "⑤ 教育公平\n(Equity)",
            "家庭社經背景（ESCS）在多大程度上決定成績？",
            "ESCS 四分位科學差距：45 分\nESCS 解釋變異：僅 4%\n學術韌性學生比例：約 19%",
            "ESCS 四分位差：85 分\nESCS 解釋變異：12%\n學術韌性學生：約 10%",
            "借鑒學術韌性學生特質，強化弱勢學生校內資源傾斜，嚴防 2015-2025 年優勢與弱勢差距擴大趨勢。"
        ],
        [
            "⑥ 學習投入\n(Engagement)",
            "學生高分背後，是否真正熱愛學習與具備歸屬感？",
            "喜歡在校學新事物：僅 58.1%\n學校歸屬感指數：偏低\n成長型思維指標：中等",
            "喜歡學新事物：68.8%\n學校歸屬感指數：中等\n成長型思維指標：中高",
            "診斷「應試高效但求知倦怠」的課堂文化，引入專案式探究學習（PBL），修復學生內在求知動機。"
        ],
        [
            "⑦ 數碼與 AI\n(Digital & AI)",
            "科技工具是賦能高階思維還是造成認知外包？",
            "每週用 AI Chatbot：67%\nAI 作業初稿：40%\nAI 文本摘要：41%\n課堂數位受干擾：13%",
            "每週用 AI：46%\nAI 作業初稿：29%\nAI 文本摘要：30%\n課堂數位受干擾：28%",
            "建立校本 AI 批判性使用守則（Q-V-R-E-R 模式）；警惕數位干擾導致科學成績重挫 24 分的嚴峻代價。"
        ],
        [
            "⑧ 學校生態\n(Ecosystem)",
            "課堂紀律、師生關係與家庭支持是否協同共振？",
            "課堂紀律氛圍優於 OECD\n教師情感與學業支援：良好\n家庭學術督導與支持：偏低",
            "課堂干擾現象普遍\n教師支援感知一般\n家庭課業溝通適中",
            "深化家校共育機制，提升家長對青春期學生自主學習的心理與情感支援，營造支持性學術生態。"
        ]
    ]
    add_styled_table(doc, headers_dim, data_dim, col_widths=[1.1, 1.3, 1.4, 1.3, 1.4])

    add_callout(doc, [
        "管理層啟示：在 8 個維度中，對推動深層課程變革最為關鍵的往往不是代表總量成績的【① 學業成就】，而是【② 分佈結構】、【③ 核心素養】、【⑤ 教育公平】、【⑥ 學習投入】與【⑦ 數碼與 AI】。",
        "若學校僅緊盯均分，將徹底掩蓋底層學生的脫節、思維能力的空洞化、學習動機的枯竭以及數位時代的認知退化。"
    ], title="架構核心價值與決策指向")

print("Part 1 built successfully.")
