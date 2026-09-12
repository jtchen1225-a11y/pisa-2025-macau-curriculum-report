# -*- coding: utf-8 -*-
"""
Part 2 of PISA 2025 School-based Curriculum Reform Report:
- Section 3: Deep Dive into Macau's 5 Core Contradictions
- Section 4: Three-Level Implementation & School Diagnostic Framework
"""

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from build_report import (
    HEX_NAVY, HEX_SLATE, HEX_DARK, HEX_MUTED, HEX_LIGHT_BG, HEX_ROW_ALT,
    HEX_ACCENT_RED, HEX_ACCENT_GOLD, HEX_WHITE, RGB_NAVY, RGB_SLATE, RGB_DARK,
    add_styled_heading, add_para, add_bullet, add_callout, add_styled_table,
    set_cell_background, set_cell_margins
)

def build_section_3(doc):
    add_styled_heading(doc, "第三章：澳門 PISA 2025 五大核心矛盾之深度剖析與校本反思", level=1)
    add_para(doc, "作為課程發展助理主任，我們絕不能被表面的耀眼均分所迷惑。深入解讀 OECD 數據庫與技術指標，澳門在本次 PISA 中浮現出五個極具張力與警示意味的「深層矛盾」。這五大矛盾恰恰指明了本校未來三至五年課程與課堂教學改革的著力點。")

    # Contradiction 1
    add_styled_heading(doc, "3.1 矛盾一：高平均分掩蓋下的「兩極化分流」隱患", level=2)
    add_para(doc, "澳門在四大領域的平均分均居世界前列（數學 549、科學 541、閱讀 501、CPS 572）。然而，在總分的光環之下，學生成就的分佈特徵卻透露出結構性隱憂：")
    add_bullet(doc, "澳門約 7.9% 的學生處於 Level 2 以下（未達基礎素養門檻），約 14.5% 達到 Level 5–6（頂尖水準）；在數學方面，達到 Level 5–6 的學生佔比高達約 29%。而在創新領域運算解難（CPS）中，更有高達 53.7% 的學生跨入 Level 5–6。", bold_prefix="成就層級分佈現狀：")
    add_bullet(doc, "2022 至 2025 年間，全澳數學最低 10%（P10）與最高 10%（P90）的差距出現擴大趨勢。這意味著在頂尖學生不斷向上突破的同時，底層學生的學習落差並未縮小，甚至正在被逐漸拉大。", bold_prefix="離散度擴大警訊：")
    add_bullet(doc, "長期以來，學校常態評量普遍依賴「班平均分 → 年級平均分 → 全校平均分」的單一線性指標。這種「均值思維（The Flaw of Averages）」是極其危險的遮羞布：一個班級平均分 75 分，可能由「高分 95 與低分 55 的極端分化」構成，也可能由「全員 70-80 分的均質狀態」構成。兩者所需的教學介入策略截然不同。", bold_prefix="校內管理層反思：")
    
    add_callout(doc, [
        "改革行動：學校教學評量儀表板（Assessment Dashboard）必須徹底告別「單一平均分考核」，全面升級為「四維分佈評量法」：",
        "① 嚴格監控低成就學生（Level <2 / 校內後 20%）比例變化；",
        "② 追蹤頂尖優異學生（Level 5–6 / 校內前 20%）的增長曲線；",
        "③ 計算並鎖定每學期段考與期末考的 P10 / P90 離散度差距；",
        "④ 建立「底線防跌、天花板拔尖」的雙向質量保證機制。"
    ], title="管理層反思：從均值看板到分佈看板", alert=False)

    # Contradiction 2
    add_styled_heading(doc, "3.2 矛盾二：數理高光 vs 閱讀滑落（-9分）——學科閱讀素養的警鐘", level=2)
    add_para(doc, "2022 至 2025 年趨勢分析中，最值得課程發展團隊警醒的數據是領域間的顯著分化：")
    add_bullet(doc, "科學微降 2 分（無統計顯著差異），數學微降 3 分（無統計顯著差異），而閱讀素養顯著下降 9 分（由 510 分驟降至 501 分）。", bold_prefix="領域趨勢對比：")
    add_bullet(doc, "學校傳統觀念普遍存在嚴重誤區，將閱讀成績下滑單純歸咎為「中文科與英文科的責任」。然而，PISA 的閱讀評估本質上是跨情境、跨媒介的認知素養，涉及長文本解碼、非連續文本（圖表、數據、示意圖）比對、多來源資訊整合及反思批判。", bold_prefix="破解學科分割迷思：")
    add_bullet(doc, "高階數理與運算解難高度依賴「學科閱讀素養（Disciplinary Literacy）」。現代 PISA 數學與科學試題極少出現單純的算式計算，均以長篇真實生活情境包裝。學生若無法讀懂長題幹、無法從冗餘資訊中篩選關鍵變量、無法理解圖表與文本的映射關係、無法用清晰語言解釋自己的推理過程（Claim-Evidence-Reasoning），數學建模將無從談起。", bold_prefix="對數學與科學課程的致命衝擊：")

    add_callout(doc, [
        "數學科課程範式轉型倡議：本校數學科必須徹底擺脫「公式講解 → 典型例題 → 機械刷題」的舊路徑，全面推動六階高階認知循環：",
        "【情境解讀】→【資訊提取與過濾】→【數學建模】→【邏輯推理】→【符號運算】→【結果解釋與現實評估】。",
        "將閱讀、理解、建模與反思深度鑲嵌於每一次單元測驗與日常課堂中，落實「全員學科閱讀（Reading Across the Curriculum）」理念。"
    ], title="學科素養轉向：數學課不是只算算術", alert=True)

    # Contradiction 3
    add_styled_heading(doc, "3.3 矛盾三：卓越公平下的「底線拉大」危機——提升上限而不拉大下限", level=2)
    add_para(doc, "澳門在 PISA 教育公平（Equity）領域一直被 OECD 譽為全球典範：")
    add_bullet(doc, "澳門最優勢（Top 25% ESCS）與最弱勢（Bottom 25% ESCS）學生的科學差距僅約 45 分，遠低於 OECD 平均的 85 分；家庭社經背景僅解釋澳門約 4% 的科學成績變異（OECD 平均約 12%）；全澳約 19% 的弱勢家庭學生在科學領域突破逆境，進入全澳前四分之一的高分群，展現出強大的「學術韌性（Academic Resilience）」。", bold_prefix="全球領先的公平性指標：")
    add_bullet(doc, "從 2015 至 2025 年的十年間，澳門最優勢與最弱勢群體的科學差距呈現微幅擴大；而 2022 至 2025 年數學高低分位數差距進一步拉開。這表明，在推動課程加深加廣、引入現代前沿科技（如編程與 AI）的過程中，文化資本與家庭支持更充足的學生獲益速度明顯快於弱勢學生，公平底色面臨侵蝕。", bold_prefix="暗流湧動的差距擴大：")

    add_callout(doc, [
        "核心戰略命題：本校課程改革絕不能重蹈「為了追求拔尖而犧牲弱勢」或「為了遷就後進而拉低學術標準」的兩難泥淖。",
        "我們未來的核心課程發展 KPI 必須定為：『如何全力提升學術上限（拔尖），同時守住且墊高學術下限（保底）？』",
        "具體抓手在於提煉那 19%「學術韌性學生」的共性特質：高自我效能感、堅定目標感、高專注度課堂參與，並將這些非認知心理資本結構化融入日常班級經營與輔導體系。"
    ], title="核心戰略命題：提升上限，守住底線")

    # Contradiction 4
    add_styled_heading(doc, "3.4 矛盾四：「高成就 × 低投入」的東亞學習文化死結", level=2)
    add_para(doc, "這可能是本次 PISA 數據對學校管理層最震撼、最需深思的反差指標：")
    add_bullet(doc, "澳門僅約 58.1% 的 15 歲學生表示「喜歡在學校學習新事物」，而 OECD 總體平均高達 68.8%（落後 OECD 逾 10 個百分點）；與此同時，澳門學生的學校歸屬感（Sense of School Belonging）與感知到的家庭學業支持亦處於相對弱勢水平。", bold_prefix="數據實證反差：")
    add_bullet(doc, "澳門學生高度擅長在結構化高壓環境下完成重複訓練與應試任務，但在離開強制監督後，往往缺乏內在求知慾望，甚至對課堂學習產生情感疏離。學生可能只是「被動高分的考試機器」，而非「熱愛探索的主動學習者」。", bold_prefix="深層文化病灶：")
    add_bullet(doc, "一旦升入高等院校或步入職場，面對無標準答案、高度模糊的真實複雜挑戰時，缺乏內在動機與「成長型思維（Growth Mindset）」的學生將迅速失去航向。若課程改革只盯著測驗分數，我們將在課堂中培養出一批「高分卻厭學」的脆弱青年。", bold_prefix="對課程發展的長期威脅：")

    add_callout(doc, [
        "課堂變革方針：由『被動灌輸型課堂』走向『探究與自主賦能型課堂』。",
        "在 P1–S6 課程中體系化導入探究式學習（Inquiry-based Learning）與真實問題解決專案（Real-world Problem Solving），賦予學生探究自主權；在評量中鼓勵試錯與迭代，打破唯分數至上的評價文化，重燃學生的好奇心與學術歸屬感。"
    ], title="重塑學習文化：高成就必須伴隨高投入", alert=True)

    # Contradiction 5
    add_styled_heading(doc, "3.5 矛盾五：高 AI 普及率下的「認知外包」危機與課堂數位干擾代價", level=2)
    add_para(doc, "PISA 2025 首次系統性調查了學生的生成式 AI（如 ChatGPT、各類 AI Chatbot）使用行為與課堂數位干擾，澳門的數據極富戲劇性：")
    
    headers_ai = ["AI 輔助學習用途", "澳門學生佔比", "OECD 平均佔比", "校本課程教學診斷"]
    data_ai = [
        ["每週至少使用 AI Chatbot 學習", "67%", "46%", "澳門學生普及度極高，AI 已深刻重構課後學習生態"],
        ["使用 AI 初步研究新課題", "38%", "31%", "正面應用：利用 AI 擴展資訊視野與先行探索"],
        ["使用 AI 摘要閱讀文本", "41%", "30%", "雙刃劍：節省時間 vs 削弱精讀、深讀與長文本耐心"],
        ["使用 AI 撰寫作業初稿", "40%", "29%", "高危領域：極易誘發認知外包，喪失獨立構思與論證能力"],
        ["幾乎從不使用 AI 完成學習任務", "5%", "約 15%", "AI 零接觸者已成極少數，禁止不如科學規範"]
    ]
    add_styled_table(doc, headers_ai, data_ai, col_widths=[1.8, 1.1, 1.1, 2.5])

    add_bullet(doc, "使用 AI 的關鍵在於區分「AI Replacing Thinking（認知外包、思維萎縮）」與「AI Supporting Thinking（思維腳手架、批判擴展）」。學生 A 遇到數學題直接將題目丟給 AI 複製答案；學生 B 則先給出自己的解題步驟，要求 AI 扮演挑剔的評審挑出推理漏洞並進行概念辯論。兩者的教育效能有雲泥之別！", bold_prefix="認知外包 vs 思維賦能：")
    add_bullet(doc, "澳門僅約 13% 的學生表示在科學課上經常受到同儕使用數位設備的干擾（OECD 為 28%），環境看似良好。然而，計量迴歸顯示：在控制社經背景後，澳門受到數位干擾的學生，其科學成績平均落後高達 24 分，而 OECD 平均差距僅 11 分！", bold_prefix="課堂數位干擾（Digital Distraction）的沉重代價：")
    add_bullet(doc, "這表明澳門課堂教學密度高、進度快，一旦學生在課堂中被平板、手機或非教學訊息分心，其認知鏈條即刻斷裂，造成的學業損耗是 OECD 國家的兩倍以上！", bold_prefix="深度剖析干擾代價：")

    add_callout(doc, [
        "學校數碼政策重構：學校的 Digital Policy 必須從過去簡單的「設備管理與配備（Device Availability）」跨越到「數位教學法 + 課堂專注度管理 + 批判性 AI 素養（Pedagogy, Discipline & Critical Literacy）」：",
        "① 課堂推行 3D 專注策略（Design 任務引導、Discipline 物理螢幕管控、Diagnostic 學習軌跡監控）；",
        "② 課程導入 Q-V-R-E-R 思維五步法，禁止未經反思的答案複製，全面引導學生將 AI 作為邏輯思辨與深度探究的對話夥伴。"
    ], title="數位與 AI 政策轉向：防範認知萎縮與課堂分心", alert=True)

def build_section_4(doc):
    add_styled_heading(doc, "第四章：三層次落地推進架構與校本課程診斷模型", level=1)
    add_para(doc, "為確保改革不淪為空洞口號，課程發展處將 PISA 洞察拆解為層層遞進的「三層分析與落地模型」：")

    # Layer 1
    add_styled_heading(doc, "4.1 第一層：東亞高表現教育體系跨域標竿比較（Macau Benchmark）", level=2)
    add_para(doc, "分析澳門無須與全球 90 個體系全部對比，而應精準錨定同處儒家文化圈、以高學業表現著稱的東亞核心標竿：澳門（MAC）、新加坡（SGP）、日本（JPN）、韓國（KOR）、中華台北（TPE）、香港（HKG）及 OECD 基準。")

    headers_bm = ["經濟體 / 指標", "數學均分", "科學均分", "閱讀均分", "運算解難(CPS)", "ESCS 解釋變異", "學習樂趣比例"]
    data_bm = [
        ["澳門 (Macau)", "549", "541", "501", "572 (全球第1)", "約 4% (極高公平)", "58.1% (偏低)"],
        ["新加坡 (Singapore)", "575+", "560+", "540+", "565+", "約 13% (高成就/高競爭)", "65% 左右"],
        ["日本 (Japan)", "535+", "545+", "515+", "550+", "約 9% (均衡發展)", "62% 左右"],
        ["中華台北 (Chinese Taipei)", "545+", "535+", "510+", "555+", "約 14% (兩極分化明顯)", "56% 左右"],
        ["韓國 (Korea)", "525+", "525+", "515+", "545+", "約 10% (補習密集)", "59% 左右"],
        ["香港 (Hong Kong)", "540+", "520+", "500+", "545+", "約 6% (波動調整)", "55% 左右"],
        ["OECD 平均基準", "463", "482", "461", "500", "12% (中等)", "68.8% (高投入)"]
    ]
    add_styled_table(doc, headers_bm, data_bm, col_widths=[1.5, 0.8, 0.8, 0.8, 1.1, 1.1, 1.0])

    add_para(doc, "標竿洞察：澳門在 CPS 運算解難與教育公平（ESCS 變異僅 4%）上傲視東亞群雄，但在「閱讀素養」與「學習樂趣投入」上，落後於新加坡與日本。新加坡兼具高閱讀與高數理，其核心經驗在於體系化推動雙語深度閱讀與探究式科學教學；本校應以新加坡為標竿，強化學科閱讀與情境論證。")

    # Layer 2
    add_styled_heading(doc, "4.2 第二層：PISA 核心認知素養能力解構（Competency Analysis）", level=2)
    add_para(doc, "PISA 評估的核心並非考核書本知識的復現，而是檢驗知識在未知情境中的動態遷移。以數學與運算解難為例：")
    add_bullet(doc, "PISA 數學素養包含三大核心認知過程：① 形成情境問題（Formulate：將現實雜亂問題轉化為數學符號與模型）；② 應用概念與技巧（Employ：執行演算法、符號推導與計算）；③ 解釋與評估（Interpret & Evaluate：將數學結論放回現實情境中檢驗合理性）。傳統課堂往往 80% 時間耗費在中間的 Employ（繁瑣計算），而 PISA 考查的精髓恰在兩端的 Formulate 與 Interpret/Evaluate！", bold_prefix="數學素養三大過程：")
    add_bullet(doc, "澳門奪冠的 CPS 領域涵蓋四大能力構念：① 問題分解（Decomposition：將龐大複雜情境拆分為子任務）；② 演算法設計（Algorithm Design：設定可執行的解難流程與邏輯規則）；③ 抽象建模（Abstraction：提煉核心變量，忽略無關噪聲）；④ 系統除錯與評估（Debugging & Evaluation：面對動態回饋及時修正策略）。這為本校跨學科科技與資訊課程提供了清晰的能力畫像。", bold_prefix="運算解難（CPS）四大能力：")

    # Layer 3
    add_styled_heading(doc, "4.3 第三層：校本課堂與評量深度審計模型（School Diagnostic）", level=2)
    add_para(doc, "外部數據最終必須精準映射回本校日常教學。課程發展處特建立六階段校本診斷與閉環審計架構：")
    add_para(doc, "【PISA Benchmark 外部標竿】→【Curriculum Mapping 課程地圖】→【School Assessment 校內測考】→【Student Performance 學生表現】→【Teaching Practice 課堂教學】→【Curriculum Reform 制度改革】", italic=True)
    
    add_para(doc, "當我們發現某一核心素養（例如『數學建模』或『跨學科圖表論證』）在校內學生表現疲軟時，管理層切忌盲目問責，而應依照「三岔路口診斷決策樹」進行根因分析：")

    headers_diag = ["審計情境", "根本原因判定", "管理本質", "精準干預處方"]
    data_diag = [
        [
            "情境 A：\n課程大綱寫了，\n但教師課堂沒教",
            "課程實施（Implementation）缺失\n教師教學進度受壓或缺乏教學法支援",
            "執行力與\n專業培訓問題",
            "重構科組備課機制，提供課堂建模示範教案與教學指引，釋放課時壓力，確保新課綱落到實處。"
        ],
        [
            "情境 B：\n教師課堂教了，\n但校內測考不考",
            "評量對齊（Assessment Alignment）脫節\n段考命題依賴舊題庫，考查與教學兩張皮",
            "命題機制與\n評價導向問題",
            "改革段考命題審查機制，強制規定情境建模與論證題型佔比（不低於 35%），倒逼評量對齊。"
        ],
        [
            "情境 C：\n課堂教了＋測考考了，\n但學生得分率仍低",
            "學習與認知（Cognitive/Pedagogy）失效\n教學方法脫離學生認知區，學科素養斷層",
            "學生理解與\n支架設計問題",
            "診斷前置概念斷層，在日常課堂中搭建思維腳手架（Scaffolding），啟動 MTSS 第二層小組輔導。"
        ]
    ]
    add_styled_table(doc, headers_diag, data_diag, col_widths=[1.5, 1.7, 1.1, 2.2])

    add_callout(doc, [
        "管理層警言：以往學校一旦看到成績不佳，往往只會下令『加強輔導、增加作業、多發練習卷』，這往往是在用無效的勤奮掩蓋制度性缺陷。",
        "只有分清是『沒教』、『沒考』還是『教了考了學生仍不會』，課程改革才能精準開刀、藥到病除！"
    ], title="審計決策邏輯：拒絕盲目加練，落實精準治理")

print("Part 2 built successfully.")
