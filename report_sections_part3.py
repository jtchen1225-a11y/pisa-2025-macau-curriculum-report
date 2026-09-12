# -*- coding: utf-8 -*-
"""
Part 3 of PISA 2025 School-based Curriculum Reform Report:
- Section 5: P1–S6 Vertical Curriculum Progression & Backward Design
- Section 6: Action Plan & 10-Dimension Implementation Matrix
- Section 7: Reform Roadmap (2026–2028)
- Section 8: Appendices, Technical Notes & References
"""

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from build_report import (
    HEX_NAVY, HEX_SLATE, HEX_DARK, HEX_MUTED, HEX_LIGHT_BG, HEX_ROW_ALT,
    HEX_ACCENT_RED, HEX_ACCENT_GOLD, HEX_WHITE, RGB_NAVY, RGB_SLATE, RGB_DARK,
    add_styled_heading, add_para, add_bullet, add_callout, add_styled_table,
    set_cell_background, set_cell_margins
)

def build_section_5(doc):
    add_styled_heading(doc, "第五章：P1–S6 數學與跨學科素養縱向貫通規劃（Vertical Progression）", level=1)
    add_para(doc, "PISA 測試的對象是 15 歲學生（初三至高一年級），但一個 15 歲少年所具備的數學抽象、跨學科閱讀、運算邏輯與論證批判能力，絕非初三年級突擊訓練所能達成，而是自小學一年級起長達九至十年認知經驗層層累積的產物。如果我們等到學生進入初三才來補救，其思維定勢與概念短板早已固化。")

    add_styled_heading(doc, "5.1 15 歲核心素養之逆向設計（Backward Design）貫通藍圖", level=2)
    add_para(doc, "課程發展處以 PISA 2025 評估框架為終點標準，採用「逆向設計（Backward Design）」理念，將 15 歲所需的高階素養向下拆解至十二年一貫的四大學習階段：")

    headers_vert = ["學習階段", "對應年級", "認知發展核心任務", "學科閱讀與建模落腳點", "AI 與運算思維融合目標"]
    data_vert = [
        [
            "第一學段：\n啟蒙具象期",
            "小學 P1–P3\n(低年級)",
            "建立直觀數感、空間幾何感知、生活數學具象操作；激發好奇心與口頭提問習慣",
            "圖畫、符號與自然語言之雙向互譯；能用完整句子口頭說明『我的算式是怎麼想出來的』",
            "非插電（Unplugged）運算遊戲；分類、排序、規律發現；圖形化指令方向控制"
        ],
        [
            "第二學段：\n過渡抽象期",
            "小學 P4–P6\n(高年級)",
            "由算術思維向初等代數思維過渡；理解分數、比例、變量；培養非連續文本圖表解讀能力",
            "多模態長題幹閱讀；提取核心條件，剔除干擾噪聲；撰寫兩至三步的解題推理步驟與驗證說明",
            "Scratch/Blockly 圖形化編程；模組化拆解（Decomposition）；條件分支與迴圈邏輯"
        ],
        [
            "第三學段：\n建模分化期\n(PISA關鍵期)",
            "初中 S1–S3\n(中一至中三)",
            "形式邏輯推理、代數幾何嚴密論證、函數與統計建模；防範學業兩極分化斷層（S2 分水嶺）",
            "學科素養閱讀（Disciplinary Literacy）；辨析 Claim-Evidence-Reasoning；真實生活情境數學建模",
            "Python 基礎文字編程；運算解難（CPS）複雜動態探索；AI 批判質疑（Q-V-R-E-R 五步法）"
        ],
        [
            "第四學段：\n深化探究期",
            "高中 S4–S6\n(中四至中六)",
            "高等數學微積分與機率統計直覺；跨學科專題研究（PBL）；複雜系統動態仿真與決策優化",
            "學術論文與科普文獻研讀；高階批判性論證與假設檢驗；獨立撰寫跨學科研究專題報告",
            "進階數據分析與演算法設計；利用 AI 進行研究構想輔助與代碼除錯；數位自主學習者定型"
        ]
    ]
    add_styled_table(doc, headers_vert, data_vert, col_widths=[1.2, 1.0, 1.5, 1.4, 1.4])

    add_styled_heading(doc, "5.2 破解中小學貫通銜接的四大「學習斷層（Learning Chasms）」", level=2)
    add_para(doc, "透過歷年校內成績大數據分析，學生在十二年學習進程中往往面臨四大關鍵斷層，這是導致後續 PISA 成績分化的主要病灶：")

    add_bullet(doc, "小學三年級以前以整數直觀計算為主，進入四年級後全面引入小數、分數、面積幾何與抽象文字題。學校必須在此階段強化教具操作與具象化視覺模型（如新加坡數學採用的 CPA 模型：Concrete-Pictorial-Abstract），防止低年級數學死記硬背導致的抽象斷崖。", bold_prefix="斷層一（P3 升 P4）：具象操作轉向符號抽象之斷層：")
    add_bullet(doc, "小學數學題幹簡短、條件封閉，進入初中後題目文字量激增、背景情境複雜，並融入跨學科資訊。初中教師常誤以為學生『小學沒學好算術』，實則是『學生讀不懂長題幹』。初一必須設置 4 至 6 週的『學科閱讀與情境破題銜接單元』。", bold_prefix="斷層二（P6 升 S1）：封閉題目轉向情境閱讀之斷層：")
    add_bullet(doc, "初二開始大量出現平面幾何嚴密推導、一次函數與反比例函數，認知要求從『直觀計算』徹底躍升為『公理化演繹』。全校歷年數學低及格率的爆發點正位於初二上學期。此處必須實施第一道 MTSS 強制干預防線，精準補底，嚴防大批學生在此階段放棄數學。", bold_prefix="斷層三（S2 升 S3）：直觀思維轉向演繹證明之斷層（最危險分水嶺）：")
    add_bullet(doc, "完成義務教育階段後，高中面臨文理分流或升學專業深化。部分學生將高階數理視為畏途。高中課程必須引入真實跨學科科研專題（如智慧城市數據建模、環境感測器數據分析），讓學生體會數學與科學作為思維工具的威力。", bold_prefix="斷層四（S3 升 S4）：普及素養轉向專業深化之斷層：")

def build_section_6(doc):
    add_styled_heading(doc, "第六章：校本課程改革行動方案與操作工具包", level=1)
    add_para(doc, "基於前述八大維度、五大矛盾與 P1–S6 逆向設計，課程發展處制訂了涵蓋行政、教學、評量與科技的十項一體化改革行動矩陣：")

    # Section 6.1: Matrix Table
    add_styled_heading(doc, "6.1 《PISA 2025 澳門 × 本校課程改革分析矩陣》", level=2)
    
    headers_mat = ["分析維度", "澳門 PISA 2025 核心發現", "本校教務與科組必問核心問題", "具體校本改革行動方案", "責任科組 / 檢驗指標"]
    data_mat = [
        [
            "Achievement\n(學業成就)",
            "數理均分全球第3，維持超高水準，但增長進入高原期",
            "本校數理優勢是否在各年級穩定保持？學科分化從何時開始？",
            "構建 P1–S6 縱向螺旋課程地圖，確保核心概念連續推進，清除學段重複與斷層。",
            "數學組 / 理科組\n年級橫向與縱向增值評量"
        ],
        [
            "Low Performers\n(低成就學生)",
            "科學 <L2 約 7.9%，數學底層分佈存在擴大隱憂",
            "校內後 20% 的學生究竟是誰？其未掌握的共通概念是什麼？",
            "全面建立 MTSS 多層級支援架構，實施每週針對性小組補底，嚴控全校學業不及格率。",
            "教務處 / 學生輔導組\n後 20% 脫困率與及格率"
        ],
        [
            "Top Performers\n(高成就學生)",
            "數學 Level 5-6 達 29%，CPS 高表現達 53.7%，尖子群龐大",
            "校內頂尖 20% 學生是否在課堂上感到吃不飽？高階題目是否足夠？",
            "設立『拔尖培優專項模組』，引進高等數學建模、奧賽訓練及跨學科專題研究（PBL）。",
            "資優教育組 / 各科組\n高階情境題掌握率 >80%"
        ],
        [
            "Reading\n(閱讀素養)",
            "顯著下滑 9 分（降至 501），成為四大領域中最脆弱短板",
            "非語文科（數理社）課堂與考卷中，學科閱讀素養是否長期被忽視？",
            "發起『全校學科閱讀素養（Disciplinary Literacy）』運動，考卷引入非連續文本與情境辨析。",
            "全校各學科組\n非連續文本命題佔比 ≥30%"
        ],
        [
            "Equity\n(教育公平)",
            "總體公平極佳（ESCS變異僅4%），但優勢與弱勢差距歷年有擴大跡象",
            "校內弱勢家庭背景學生是否在進階選修與科技活動中處於邊緣？",
            "實施弱勢學生學術扶持計劃，提供免費校內課後數位學習資源與學業導師配對。",
            "教務處 / 總務處\n弱勢學生高表現進展追蹤"
        ],
        [
            "CPS\n(運算解難)",
            "平均 572 分全球第一，展現運算工具解難之驚人潛力",
            "常態課程是否常態化培養學生的演算法思維、抽象建模與系統除錯？",
            "從小學三年級起普及圖形化與文字編程，將運算思維工具嵌入數學與科學專題探究。",
            "資訊科技組 / 跨學科組\n全校編程與CPS專案普及率"
        ],
        [
            "Engagement\n(學習投入)",
            "僅 58.1% 喜歡在校學新事物（落後 OECD 逾 10%），學習歸屬感低",
            "高分學生是否真正熱愛所學？課堂是否充滿枯燥的被動刷題？",
            "推動課堂探究教學法，增加情境體驗、動手實驗與學生自主選題，重構學習內在驅動力。",
            "教研組 / 班主任團隊\n學生課堂投入度與滿意度"
        ],
        [
            "AI Pedagogy\n(人工智慧)",
            "67% 每週使用 AI，40% 用於作業初稿，存在高度認知外包風險",
            "學生是在利用 AI 取代思考，還是利用 AI 拓展深度思維？",
            "制訂校本生成式 AI 學習指引，推行 Q-V-R-E-R 批判思考五步法，禁止未標註之答案搬運。",
            "課程發展處 / 資訊科技組\nAI 批判性使用作業認證"
        ],
        [
            "Family\n(家庭支持)",
            "家庭學術溝通與情感支持指標相對落後",
            "家長是否只關心排名分數，忽視了心理韌性與探索興趣的培養？",
            "舉辦『現代素養與家庭共育工作坊』，指導家長從『催促分數』轉向『傾聽與探究引導』。",
            "家校協同委員會\n家長工作坊參與度與回饋"
        ],
        [
            "Digital Policy\n(數位管理)",
            "課堂受數位干擾學生科學重挫 24 分，負面衝擊極具破壞性",
            "校內課堂平板/電腦使用是否存在注意力渙散與隱蔽玩樂？",
            "實施課堂數位專注『3D 策略』（任務設計、螢幕監管、軌跡診斷），杜絕非學習性干擾。",
            "教導處 / 全體科任教師\n課堂專注度與設備違規率"
        ]
    ]
    add_styled_table(doc, headers_mat, data_mat, col_widths=[1.1, 1.3, 1.4, 1.5, 1.2])

    # Section 6.2: Assessment Dashboard Upgrade
    add_styled_heading(doc, "6.2 校內教學評量儀表板（Assessment Dashboard）升級方案", level=2)
    add_para(doc, "傳統成績單只呈現總分與全級排名，無法為課程改善提供任何診斷資訊。教務處將自 2026/2027 學年第二學期起，全面上線新型『素養導向教學評量儀表板』：")
    add_bullet(doc, "以盒鬚圖（Boxplot）呈現各班成績分佈，自動標出 P10（低分底線）、P50（中位數）與 P90（高分天花板）。考核指標以『P10 向上躍升幅度』與『四分位距（IQR）收斂度』為核心，嚴防兩極分化。", bold_prefix="指標一：P10 / P90 離散度監控：")
    add_bullet(doc, "建立全校學業紅黃綠燈預警系統，凡單元評量連續兩次落入後 20%（相當於 PISA Level <2）者，系統自動觸發預警並推送個人化補底任務單。", bold_prefix="指標二：低成就門檻（Level <2）預警機制：")
    add_bullet(doc, "段考與期末考命題嚴格執行『情境化、開放性、論證性試題佔比不低於 35%』的硬性規定。每道情境大題需明確標註考查的認知維度（Formulate、Employ 或 Interpret），建立題庫素養雙向細目表。", bold_prefix="指標三：試卷情境題與學科閱讀權重審查：")

    # Section 6.3: MTSS Framework
    add_styled_heading(doc, "6.3 MTSS（多層級支援架構）實施細則：提升上限，守住底線", level=2)
    add_para(doc, "為破解『拔尖 vs 保底』的二元對立，學校引進國際通行的 MTSS（Multi-Tiered System of Supports）三層加一層干預架構：")
    add_bullet(doc, "覆蓋全校 100% 學生。全面落實『通用設計學習法（Universal Design for Learning, UDL）』，課堂提供多種資訊呈現方式（文本、圖表、動態模擬）與多樣化表達途徑，確保 80% 以上學生在常態課堂完成基礎素養達標。", bold_prefix="Tier 1（第一層：普及性高質量常態課堂）：")
    add_bullet(doc, "針對全校後 15% 至 20% 處於學業邊緣的學生。每週開設 2 節固定時段的『學科閱讀與建模思維強化小組』。以 4–6 人微型小組開展精準支架教學，重點攻克概念盲區與題幹閱讀障礙，阻斷分化。", bold_prefix="Tier 2（第二層：針對性小組靶向補強）：")
    add_bullet(doc, "針對後 3% 至 5% 存在嚴重學習困難的學生。啟動駐校特教老師、心理輔導員與學科名師的一對一個別化教育方案（IEP），進行深度認知診斷與心理減壓，提供定製化支持。", bold_prefix="Tier 3（第三層：個別化深度專業介入）：")
    add_bullet(doc, "針對前 15% 至 20% 學有餘力的拔尖學生。開設『高階數理探究』、『跨學科 AI 與科研創新專案』及頂尖賽事工作坊，由校內科研導師帶領進行開放式研究，激發學術潛力。", bold_prefix="Enrichment Tier（第四層：卓越拔尖培優機制）：")

    # Section 6.4: AI Q-V-R-E-R Framework
    add_styled_heading(doc, "6.4 校本「AI 賦能課堂與批判性思維五步法」（Q-V-R-E-R 模式）", level=2)
    add_para(doc, "面對澳門 67% 的極高 AI 使用率，單純禁止如同築堤擋水。課程發展處特為全校師生制定『Q-V-R-E-R 批判思考五步法』，要求全校專題作業與課後探究必須遵循此範式：")
    add_bullet(doc, "學生不直接向 AI 索取最終答案，而是學習將複雜現實問題拆解為具備清晰約束條件、背景設定與邏輯邊界的結構化提示詞（Prompts）。", bold_prefix="步驟一：Questioning（精確提問與情境拆解）：")
    add_bullet(doc, "嚴格禁止盲信 AI 輸出。學生必須運用權威教材、學術數據庫或物理實驗，對 AI 提供的數據、引文與事實進行跨源交叉查核（Cross-Verification），尋找幻覺（Hallucination）漏洞。", bold_prefix="步驟二：Verification（跨源交叉驗證）：")
    add_bullet(doc, "學生需關閉 AI 視窗，用自己的邏輯語言手寫或重述解題推理路徑（Human-in-the-loop），對比自身思路與 AI 演算路徑的優劣，強化大腦內部的神經突觸連接。", bold_prefix="步驟三：Reasoning（獨立推理與思維對照）：")
    add_bullet(doc, "批判性評估 AI 解法的效率、局限性、假設前提與倫理邊界。追問：『這個模型適用於所有極端情況嗎？有沒有更簡潔的思維模型？』", bold_prefix="步驟四：Evaluation（批判評估與邊界探討）：")
    add_bullet(doc, "在作業末尾撰寫 150 字的『AI 協作反思日記』，詳述：我向 AI 提出了什麼質疑？AI 哪裡犯了錯？我在這場對話中學到了什麼元認知策略？", bold_prefix="步驟五：Reflection（元認知反思與知識內化）：")

    # Section 6.5: Digital Distraction Guidelines
    add_styled_heading(doc, "6.5 課堂數位設備防干擾管理守則（3D 專注策略）", level=2)
    add_para(doc, "針對澳門受到課堂數位干擾學生科學重挫 24 分的慘痛教訓，學校嚴禁『無目的、無約束的設備自由使用』，全面實施 3D 專注策略：")
    add_bullet(doc, "課堂使用平板或電腦必須對應明確的『微任務單（Micro-tasks）』，每次數位操作時間控制在 10–15 分鐘以內，任務完成即刻切換回板書與同儕研討，杜絕長時間盯屏導致的思維游離。", bold_prefix="① Design（任務驅動設計）：")
    add_bullet(doc, "制定嚴格的『螢幕可見原則（Screens Up / Screens Down）』。教師下達指令時全員合上螢幕；教導處部署課堂集中管理軟體，屏蔽非教學網站與社群推播，實施課堂物理與技術雙重專注防護。", bold_prefix="② Discipline（課堂紀律剛性約束）：")
    add_bullet(doc, "利用智慧校園後台大數據，監控各班設備使用時長與課堂互動日誌。對設備使用率過高但成績下滑的班級進行個別化教學法診斷與課堂視導。", bold_prefix="③ Diagnostic（軌跡監控與學情診斷）：")

    # Section 6.6: Growth Mindset
    add_styled_heading(doc, "6.6 學習文化重塑：培育成長型思維（Growth Mindset）與學術歸屬感", level=2)
    add_para(doc, "針對『高成就 × 低投入（僅 58.1% 喜歡在校學新事物）』的心智危機，學校將採取三大文化重塑工程：")
    add_bullet(doc, "課堂大力表彰『富有創見的錯誤（Productive Failure）』。將錯題分析轉化為思維迭代的勳章，建立『失敗博物館』或『錯題反思牆』，讓學生深刻體會『大腦像肌肉一樣，在克服困難中不斷生長』。", bold_prefix="① 重塑試錯文化：")
    add_bullet(doc, "每學期安排 2 週的『真實問題解難週（Real-world Problem Week）』，打破學科藩籬，讓學生圍繞澳門本地可持續發展、智慧交通、綠色能源等課題進行團隊跨學科調研，體驗所學知識造福社區的深層價值感。", bold_prefix="② 引入情境化真實探究：")
    add_bullet(doc, "構建『同儕學術導師制（Peer Tutoring）』與『師生思維午餐會』，增強師生之間的非評量性學術交流，切實提升學生在校園中的心理安全感與學術歸屬感。", bold_prefix="③ 強化校園情感與學術歸屬：")

def build_section_7(doc):
    add_styled_heading(doc, "第七章：推進時間表與工作路徑圖（2026–2028）", level=1)
    add_para(doc, "本項課程與評量深層改革是一項系統性工程，規劃自 2026 年秋季啟動，分四個階段歷時兩年半全面落實：")

    headers_rd = ["推進階段", "時間區間", "核心攻堅任務", "關鍵交付產物與檢驗節點"]
    data_rd = [
        [
            "第一階段：\n診斷與標準建立期",
            "2026 年 10 月\n至 12 月\n(2026/27 上學期)",
            "成立『校本 PISA 課程智慧專案小組』；完成全校近三年段考與期末考題目的素養審計；發布《校本學科閱讀與情境命題指引》。",
            "① 全校段考試卷審計報告\n② 新版命題雙向細目表\n③ MTSS 後 20% 預警名單"
        ],
        [
            "第二階段：\n試點推行與實驗期",
            "2027 年 1 月\n至 6 月\n(2026/27 下學期)",
            "在小四（P4）與初一（S1）先行試點『學科閱讀銜接單元』；在初二（S2）全面試行 MTSS 二級輔導小組；在全校推行 AI Q-V-R-E-R 作業模式。",
            "① P4/S1 試點示範教案集\n② AI 批判性作業範例庫\n③ 數位專注 3D 課堂視導記錄"
        ],
        [
            "第三階段：\n全面貫通與深水期",
            "2027 年 9 月\n至 2028 年 1 月\n(2027/28 上學期)",
            "正式上線校內『四維分佈教學評量儀表板』；完成 P1–S6 數學與理科縱向螺旋課程大綱修訂；舉辦首屆『校本真實情境解難展示週』。",
            "① P1–S6 縱向課程地圖手冊\n② 新型評量看板全面投產\n③ 學生學習樂趣問卷後測"
        ],
        [
            "第四階段：\n成效複盤與固化期",
            "2028 年 2 月\n至 7 月\n(2027/28 下學期)",
            "進行全校第一輪素養改革增值評估；檢驗 P10/P90 離散度與低成就比例改善成效；提煉校本成功案例，形成長效教學管理機制。",
            "① 課程改革兩年成效評估報告\n② P10 提升與及格率達標驗收\n③ 優秀教學法校本論文集"
        ]
    ]
    add_styled_table(doc, headers_rd, data_rd, col_widths=[1.2, 1.2, 2.3, 1.8])

def build_section_8(doc):
    add_styled_heading(doc, "第八章：參考文獻、權威數據庫入口與技術術語表", level=1)
    
    add_styled_heading(doc, "8.1 官方權威數據庫與文獻入口", level=2)
    add_bullet(doc, "OECD (2026). PISA 2025 Results: Learning in the Digital World & Country Notes - Macao (China). OECD Publishing, Paris. https://www.oecd.org/en/publications/pisa-2025-results-volume-i-country-notes_2d4ff9ea-en/macau-china_9815d707-en.html", bold_prefix="[1] OECD Country Note: ")
    add_bullet(doc, "OECD Education GPS (2026). Student Performance, Equity, AI & Learning Environment: Macao. https://gpseducation.oecd.org/CountryProfile?primaryCountry=MAC&topic=PI&treshold=5", bold_prefix="[2] OECD Education GPS: ")
    add_bullet(doc, "OECD (2026). PISA 2025 Database: Public Use Files (PUFs), Questionnaires & Codebooks. https://www.oecd.org/en/data/datasets/pisa-2025-database.html", bold_prefix="[3] OECD PISA Database: ")
    add_bullet(doc, "澳門特別行政區政府教育及青年發展局 (2026). 學生能力國際評估計劃（PISA 2025）研究結果新聞發佈會官方公報. https://www.gov.mo/zh-hans/news/934453/", bold_prefix="[4] 澳門教青局官方公報: ")
    add_bullet(doc, "OECD (2025). PISA 2025 Science & Learning in the Digital World Strategic Frameworks. OECD Education Working Papers.", bold_prefix="[5] PISA 2025 測評框架: ")

    add_styled_heading(doc, "8.2 核心專業術語與計量概念解析", level=2)
    add_bullet(doc, "由 OECD 根據學生家庭耐用品、父母受教育程度及職業地位合成的標準化連續變量，均值為 0，標準差為 1。用以衡量學生家庭的文化與經濟資本。", bold_prefix="● ESCS（經濟、社會與文化地位指數）：")
    add_bullet(doc, "指處於本經濟體 ESCS 最不利的四分之一（Bottom 25%），但在 PISA 評估中成績達到該經濟體前四分之一（Top 25%）的逆境自強學生。澳門達 19%，顯著優於 OECD 平均。", bold_prefix="● 学术韧性学生（Academically Resilient Students）：")
    add_bullet(doc, "PISA 不採用單一確定測驗分數，而是透過項目反應理論（IRT）與背景變量潛在回歸，為每位學生推導出的一組可能能力值（通常為 10 個）。嚴禁直接取均值跑簡單迴歸，需使用多重插補合併法估算。", bold_prefix="● 合理值（Plausible Values, PVs）：")
    add_bullet(doc, "多層級支援架構（Multi-Tiered System of Supports），結合循證教學（Tier 1）、小組精準靶向干預（Tier 2）與個別化深層介入（Tier 3），實現全體學生的差異化成功。", bold_prefix="● MTSS 多層級支援系統：")
    add_bullet(doc, "非一般基礎語文閱讀，而是指在特定學科（如數學、物理、化學）內部，解碼特定符號、圖表、論證架構及專業語言的高階認知理解能力。", bold_prefix="● 學科素養閱讀（Disciplinary Literacy）：")

    # Document Close Note
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(16)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sign = p_end.add_run("報告起草：學校課程發展處 課程發展助理主任\n審閱核定：學校教務委員會 · 課程發展專案小組\n2026 年 9 月 11 日")
    r_sign.font.name = "微軟正黑體"
    r_sign.font.size = Pt(9.5)
    r_sign.font.color.rgb = RGBColor(113, 128, 150)
    r_sign.bold = True

print("Part 3 built successfully.")
