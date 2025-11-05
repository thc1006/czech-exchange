#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交換經驗分享會簡報製作腳本 - 完全重新設計版本
講者：蔡秀吉
對象：百川學士學位學程同學
尺寸：16:9 (13.333" x 7.5")
設計理念：現代、專業、簡潔、置中對齊
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def create_presentation():
    # 創建簡報物件 (16:9)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 專業配色方案
    COLOR_PRIMARY = RGBColor(30, 58, 138)      # 深藍色主色
    COLOR_SECONDARY = RGBColor(59, 130, 246)   # 亮藍色
    COLOR_ACCENT = RGBColor(239, 68, 68)       # 紅色強調
    COLOR_DARK = RGBColor(15, 23, 42)          # 深灰文字
    COLOR_LIGHT = RGBColor(241, 245, 249)      # 淺灰背景
    COLOR_WHITE = RGBColor(255, 255, 255)      # 白色

    def add_cover_slide(title, subtitle):
        """封面投影片 - 全新設計"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 背景漸層效果（使用深藍色背景）
        background = slide.shapes.add_shape(
            1,  # Rectangle
            Inches(0), Inches(0),
            Inches(13.333), Inches(7.5)
        )
        background.fill.solid()
        background.fill.fore_color.rgb = COLOR_PRIMARY
        background.line.fill.background()

        # 主標題
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(2.5),
            Inches(11.333), Inches(1.5)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.font.size = Pt(54)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

        # 副標題
        subtitle_box = slide.shapes.add_textbox(
            Inches(1), Inches(4.2),
            Inches(11.333), Inches(0.8)
        )
        tf = subtitle_box.text_frame
        tf.text = subtitle
        p = tf.paragraphs[0]
        p.font.size = Pt(28)
        p.font.color.rgb = COLOR_SECONDARY
        p.alignment = PP_ALIGN.CENTER

        # 底部裝飾線
        line = slide.shapes.add_shape(
            1,
            Inches(4.667), Inches(5.5),
            Inches(4), Inches(0.08)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_ACCENT
        line.line.fill.background()

        return slide

    def add_section_title_slide(title, subtitle=""):
        """章節標題投影片"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 左側色塊
        left_block = slide.shapes.add_shape(
            1,
            Inches(0), Inches(0),
            Inches(5), Inches(7.5)
        )
        left_block.fill.solid()
        left_block.fill.fore_color.rgb = COLOR_PRIMARY
        left_block.line.fill.background()

        # 標題（置中）
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(3),
            Inches(11.333), Inches(1.5)
        )
        tf = title_box.text_frame
        tf.text = title
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

        if subtitle:
            subtitle_box = slide.shapes.add_textbox(
                Inches(1), Inches(4.5),
                Inches(11.333), Inches(0.6)
            )
            tf = subtitle_box.text_frame
            tf.text = subtitle
            p = tf.paragraphs[0]
            p.font.size = Pt(20)
            p.font.color.rgb = COLOR_SECONDARY
            p.alignment = PP_ALIGN.CENTER

        return slide

    def add_content_slide_centered(title, content_items, highlight_indices=None):
        """內容投影片 - 完全置中設計"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 標題區域
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.6),
            Inches(11.733), Inches(0.9)
        )
        tf = title_box.text_frame
        tf.text = title
        p = tf.paragraphs[0]
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY
        p.alignment = PP_ALIGN.CENTER

        # 標題底線
        line = slide.shapes.add_shape(
            1,
            Inches(5.167), Inches(1.5),
            Inches(3), Inches(0.05)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_SECONDARY
        line.line.fill.background()

        # 內容區域 - 置中對齊
        content_box = slide.shapes.add_textbox(
            Inches(1.5), Inches(2.2),
            Inches(10.333), Inches(4.8)
        )
        tf = content_box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP

        highlight_indices = highlight_indices or []

        for i, item in enumerate(content_items):
            if i > 0:
                tf.add_paragraph()

            p = tf.paragraphs[i]
            p.text = item
            p.font.size = Pt(20)
            p.space_before = Pt(10) if i > 0 else Pt(0)
            p.space_after = Pt(10)
            p.line_spacing = 1.4

            # 強調項目
            if i in highlight_indices:
                p.font.bold = True
                p.font.color.rgb = COLOR_ACCENT
                p.font.size = Pt(22)
            else:
                p.font.color.rgb = COLOR_DARK

            # 判斷是否為標題行（含【】或以粗體顯示）
            if item.startswith("【") or item.startswith("▸"):
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY
                p.font.size = Pt(22)

        return slide

    def add_two_column_slide_centered(title, left_content, right_content,
                                     left_title="", right_title=""):
        """雙欄投影片 - 完全置中設計"""
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 標題
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.5),
            Inches(11.733), Inches(0.8)
        )
        tf = title_box.text_frame
        tf.text = title
        p = tf.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY
        p.alignment = PP_ALIGN.CENTER

        # 中間分隔線
        divider = slide.shapes.add_shape(
            1,
            Inches(6.667), Inches(1.5),
            Inches(0.03), Inches(5.5)
        )
        divider.fill.solid()
        divider.fill.fore_color.rgb = COLOR_LIGHT
        divider.line.fill.background()

        # 左欄標題
        if left_title:
            left_title_box = slide.shapes.add_textbox(
                Inches(0.8), Inches(1.5),
                Inches(5.5), Inches(0.5)
            )
            tf = left_title_box.text_frame
            tf.text = left_title
            p = tf.paragraphs[0]
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = COLOR_SECONDARY
            p.alignment = PP_ALIGN.CENTER

        # 左欄內容
        left_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(2.2) if left_title else Inches(1.7),
            Inches(5.5), Inches(5)
        )
        tf = left_box.text_frame
        tf.word_wrap = True

        for i, item in enumerate(left_content):
            if i > 0:
                tf.add_paragraph()
            p = tf.paragraphs[i]
            p.text = item
            p.font.size = Pt(18)
            p.font.color.rgb = COLOR_DARK
            p.space_before = Pt(8)
            p.line_spacing = 1.3

            if item.startswith("【"):
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY
                p.font.size = Pt(20)

        # 右欄標題
        if right_title:
            right_title_box = slide.shapes.add_textbox(
                Inches(7.033), Inches(1.5),
                Inches(5.5), Inches(0.5)
            )
            tf = right_title_box.text_frame
            tf.text = right_title
            p = tf.paragraphs[0]
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = COLOR_ACCENT
            p.alignment = PP_ALIGN.CENTER

        # 右欄內容
        right_box = slide.shapes.add_textbox(
            Inches(7.033), Inches(2.2) if right_title else Inches(1.7),
            Inches(5.5), Inches(5)
        )
        tf = right_box.text_frame
        tf.word_wrap = True

        for i, item in enumerate(right_content):
            if i > 0:
                tf.add_paragraph()
            p = tf.paragraphs[i]
            p.text = item
            p.font.size = Pt(18)
            p.font.color.rgb = COLOR_DARK
            p.space_before = Pt(8)
            p.line_spacing = 1.3

            if item.startswith("【"):
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY
                p.font.size = Pt(20)

        return slide

    # ============ 開始製作投影片 ============

    # 投影片 1: 封面
    add_cover_slide(
        "交換不只是出國玩",
        "給任何交換生都適用的生存指南 | 蔡秀吉 @ 捷克交換經驗分享"
    )

    # 投影片 2: 章節 - 為什麼要交換
    add_section_title_slide(
        "為什麼百川人更需要出國交換？",
        "Understanding Your Unique Position"
    )

    # 投影片 3: 百川人的挑戰與機會
    add_content_slide_centered(
        "百川人的獨特處境",
        [
            "身為百川人，校級交換管道受限",
            "",
            "需自行開闢外軌途徑（如教育部獎學金）",
            "",
            "跨領域背景 = 國際競爭力的優勢",
            "",
            "全球化視野對百川人的職涯發展至關重要",
            "",
            "【核心挑戰】如何在有限資源下，找到屬於自己的路？"
        ],
        [8]
    )

    # 投影片 4: 時間軸
    add_content_slide_centered(
        "交換準備時間軸：提早規劃是王道",
        [
            "▸ 12-18個月前：決定目標、開始研究國家/學校",
            "",
            "▸ 10-12個月前：準備英文考試（托福/雅思）、提升GPA",
            "",
            "▸ 8-10個月前：撰寫研究計畫、爭取教授邀請函",
            "",
            "▸ 6-8個月前：申請獎學金（學海計畫等）",
            "",
            "▸ 4-6個月前：申請學校、準備申請文件",
            "",
            "▸ 2-4個月前：申請簽證（德國需2個月！）",
            "",
            "▸ 1-2個月前：訂機票、保險、住宿、網卡",
            "",
            "【重要提醒】越早開始，選擇越多！"
        ],
        [14]
    )

    # 投影片 5: 釐清目的
    add_two_column_slide_centered(
        "第一個決定：你為什麼要交換？",
        [
            "【玩樂導向】",
            "• 選交通樞紐城市",
            "• 有國際機場",
            "• 低成本航空發達",
            "• 例：布拉格、阿姆斯特丹",
            "",
            "【平衡型】",
            "• 學期中專注學習",
            "• 假期密集旅遊",
            "• 一週上4天課仍能遊17國",
            "  （真實案例）"
        ],
        [
            "【進修導向】",
            "• 選研究資源豐富的學校",
            "• 看教授專長是否匹配",
            "• 考慮未來就業地緣關係",
            "• 例：波士頓（生技）、矽谷（科技）",
            "",
            "【重要提醒】",
            "• 不要只學一套快淘汰的技術",
            "• 技術迭代快，重視思維與方法",
            "• 了解交換原因才能最大化收穫"
        ],
        "類型一：玩樂優先",
        "類型二：學術優先"
    )

    # 投影片 6: 申請三大支柱
    add_content_slide_centered(
        "申請成功的三大支柱",
        [
            "【支柱一】英文能力（最重要！可提前準備）",
            "   • 多數學校要求：托福80+ / 雅思6.0+ / 在校英文80分+",
            "   • 建議達到B2等級",
            "   • 越早考越好，可以重考",
            "",
            "【支柱二】在校成績 GPA（可量化指標）",
            "   • 沒有學術發表時，GPA是審核第一關",
            "   • 從大二就要開始重視",
            "",
            "【支柱三】動機與研究計畫（展現熱忱）",
            "   • 展現高度動機與熱忱",
            "   • 計畫要有組織性、可行性",
            "   • 證明你知道如何運用該地資源",
            "   • 找同學/老師進行同儕審查（peer review）"
        ]
    )

    # 投影片 7: 獎學金地圖
    add_content_slide_centered(
        "獎學金地圖：不要錯過免費機會",
        [
            "【選項A】教育部學海計畫（大學交換）",
            "   • 學海飛揚/惜珠：交換讀書類，NT$ 5-30萬",
            "   • 學海逐夢：實習類",
            "   • 弱勢優秀生（清寒）有保障名額",
            "",
            "【選項B】赴捷克短期進修獎學金（講者案例）",
            "   • 每月12,000捷克克朗（約NT$ 16,800）",
            "   • 取得邀請函 ≈ 保送",
            "   • 學校推薦名額有限（陽明交大最多5個）",
            "   • 注意：台灣錄取 ≠ 捷克最終核准",
            "",
            "【選項C】其他管道",
            "   • 教育部留學獎學金（碩博士）",
            "   • 教育部就學貸款（碩士100萬、博士200萬，清寒免息）"
        ]
    )

    # 投影片 8: 目的地選擇
    add_content_slide_centered(
        "如何選擇交換目的地？",
        [
            "【考量因素清單】",
            "",
            "• 生活成本（東歐 < 西歐 < 北美）",
            "• 交通便利性（旅遊需求）",
            "• 學術資源（進修需求）",
            "• 語言障礙程度（英語普及度）",
            "• 行政效率（歐洲普遍較慢，需心理準備）",
            "• 地緣關係（未來就業考量）",
            "• 氣候與日照（11月後歐洲4點天黑，易憂鬱）",
            "• 簽證難度（德國2個月 vs 日本3天）",
            "",
            "【特殊領域提醒】",
            "AI研究：歐盟法規嚴格（GDPR、AI法），工具受限",
            "建議在歐盟做「只能在歐盟做的事」（如網路治理、AI法案研究）"
        ]
    )

    # 投影片 9: 行前準備
    add_two_column_slide_centered(
        "行前準備：不可妥協的8件事",
        [
            "【文件類】",
            "• 簽證（提早2-3個月辦理）",
            "• 護照效期（需超過6個月）",
            "• 國際學生證 ISIC（歐洲很好用）",
            "• 英文版成績單、在學證明",
            "",
            "【財務類】",
            "• 至少2張信用卡（備援）",
            "• 了解ATM提款機制",
            "• 換一些現金在身上",
            "• 保險（醫療費用超貴！）"
        ],
        [
            "【生活類】",
            "• 網卡/SIM卡（必備！）",
            "• 翻譯軟體（非英語系國家）",
            "• 正裝西裝（歐洲考試/活動需要）",
            "• 消費電子產品在台灣買（歐洲VAT 15-19%）",
            "",
            "【行政類】",
            "• 確認選課與住宿",
            "• 查詢國定假日（避免落地遇假期）",
            "• 了解「銀行週」時間",
            "• 加入當地留學生社團（FB/Line）"
        ],
        "必備項目（左）",
        "實用工具（右）"
    )

    # 投影片 10: 財務智慧
    add_content_slide_centered(
        "財務智慧：省錢策略與陷阱",
        [
            "【省錢策略】",
            "• 歐洲VAT高（15-19%），貴重3C在台灣買（相機、耳機、GoPro）",
            "• 使用AI工具時掛VPN回台灣刷卡（避免歐盟VAT）",
            "• 了解各國退稅門檻與流程",
            "• 學會煮飯（外食比台北貴）",
            "• 購買月票/季票（德國49歐月票超划算）",
            "",
            "【常見陷阱】",
            "• 一次付清代辦費（建議分期，信用卡付款可追回）",
            "• 沒買保險（國外醫療費天價）",
            "• 護照「Republic of China」問題（建議遮China或強調Taiwan）",
            "• 被當中國人收保證金（辦電信時）",
            "",
            "【申訴管道】台灣1950 / 國外找駐外機構"
        ]
    )

    # 投影片 11: 抵達後48小時
    add_content_slide_centered(
        "抵達後48小時：生存檢查清單",
        [
            "【第一天必做】",
            "• 調時差（立刻調整作息）",
            "• 確認住宿（檢查設施、記錄損壞）",
            "• 購買網卡/申辦門號",
            "• 買基本食物、日用品",
            "",
            "【48小時內完成】",
            "• 開立當地銀行帳戶（需護照、入學證明、住址證明）",
            "• 辦理學生證（各種優惠的關鍵）",
            "• 購買交通月票",
            "• 確認教室位置（提前踩點！門牌難找）",
            "• 日本特別注意：到區公所辦轉入手續",
            "",
            "【重要提醒】歐洲行政效率慢，辦事可能需3個工作日",
            "            建議在「銀行週」前抵達，一次辦完所有行政事務"
        ]
    )

    # 投影片 12: 章節 - 軟實力
    add_section_title_slide(
        "軟實力：文化適應與溝通",
        "Soft Skills for Success"
    )

    # 投影片 13: 文化衝擊
    add_two_column_slide_centered(
        "文化衝擊：預期 vs 現實",
        [
            "【正面文化差異】",
            "• 學生支援完善（國際學生社團）",
            "• Orientation Week 很實用",
            "• 心理諮商資源充足",
            "• 特教資源多元（ADHD、資優）",
            "• 法律/宗教/藝術治療諮商",
            "",
            "【學術差異】",
            "• 老師「放牛吃草」（需自立自強）",
            "• 重視自主學習能力",
            "• 脈絡思考要嚴謹",
            "• 補考機會多"
        ],
        [
            "【可能遇到的挑戰】",
            "• 行政效率超慢（3天處理一件事）",
            "• 亞裔歧視（被查票機率高）",
            "• 語言障礙（即使在英語國家）",
            "• 季節性憂鬱（11月4點天黑）",
            "• 宿舍條件差（跳蚤、設備壞）",
            "• 校園、街上普遍抽菸",
            "",
            "【應對策略】",
            "• 不要將台灣預期套用異地",
            "• 保持禮貌，讓對方產生虧欠感",
            "• 提前修跨文化溝通課程",
            "• 服用退黑激素調節（需懂生理學）"
        ],
        "優勢面向",
        "挑戰面向"
    )

    # 投影片 14: 溝通技巧
    add_content_slide_centered(
        "溝通技巧：Speak Up & Work Smart",
        [
            "【Speak Up - 為自己發聲】",
            "• 在會議中勇於表達意見（否則會錯失資源）",
            "• 清楚表達：問題 + 解決方案 + 所需資源",
            "• 記住：大家母語都不是英文，別怕開口",
            "",
            "【破冰神器】",
            "• 食物與音樂 = 人類接受度最高的事物",
            "• 準備一些台灣小吃/音樂介紹",
            "",
            "【Work Smart, Not Just Work Hard】",
            "• 將精力花在「對未來有影響力」的事上",
            "• 找到個人利基（personal niche）",
            "• 例：不擅長光學但擅長養細胞 → 以此合作",
            "",
            "【文化差異理解】",
            "• 台灣 = 高語境文化（high context）+ 集體主義",
            "• 歐美 = 低語境文化 + 個人主義"
        ]
    )

    # 投影片 15: 心態調整
    add_content_slide_centered(
        "心態調整：從舒適圈到成長區",
        [
            "【設定明確目標】",
            "• 給自己一個目標（如周遊某地區）",
            "• 督促自己安排課業與旅行",
            "• 避免浪費假期在宿舍",
            "",
            "【Try New Things】",
            "• 交換 = 新環境 + 新開始",
            "• 嘗試以前不敢做的事（獨自旅行、看體育賽事）",
            "• 解決困難的過程 → 建立自信",
            "",
            "【找到研究/學習的熱情】",
            "• 研究突破的喜悅 > 玩樂的快樂",
            "• 避免「玩物喪志」",
            "• 研究90%無聊，但突破那10%很值得",
            "",
            "【代表學校的責任】你代表台科大、代表台灣，把自己做好才能贏得尊重"
        ]
    )

    # 投影片 16: 時間管理
    add_content_slide_centered(
        "時間管理：讀書與旅遊的平衡術",
        [
            "【核心原則】讀書第一，旅遊第二",
            "",
            "可行的平衡模式：",
            "• 一週上4天課，仍能旅遊17個國家（真實案例）",
            "• 週末短途旅行（低成本航空發達）",
            "• 學期中密集上課，假期密集旅遊",
            "• 提前在台灣列好想去的地點清單",
            "",
            "避免的陷阱：",
            "• 完全不旅遊（浪費地理優勢）",
            "• 過度旅遊（學業掛掉）",
            "• 臨時起意（機票住宿貴、時間浪費）",
            "",
            "時間管理技巧：",
            "• 制定8天短期目標（研究/學習壓力大時）",
            "• 但要注意：11月後日照短，易心態炸裂"
        ]
    )

    # 投影片 17: 給百川同學的建議
    add_content_slide_centered(
        "給百川同學的特別建議",
        [
            "【百川人的獨特優勢】",
            "• 跨領域背景 = 國際競爭力",
            "• 能整合不同領域的視角",
            "• 適合申請需要創新思維的計畫",
            "",
            "【開闢外軌途徑】",
            "1. 教育部獎學金（不限姊妹校）",
            "2. 自己聯繫國外教授爭取邀請函",
            "3. 大專生研究計畫可帶到國外做",
            "4. 參加國際競賽（如太空競賽）拓展人脈",
            "",
            "【大二/大三就要開始準備】",
            "• 準備英文版履歷、提升GPA、考英文檢定、找指導教授討論計畫",
            "",
            "【備註】大五/畢業生也能申請（透過學校推薦）"
        ]
    )

    # 投影片 18: 核心要點總結
    add_content_slide_centered(
        "核心要點總結",
        [
            "▸ 提早規劃（12-18個月前開始）",
            "",
            "▸ 釐清目的（玩樂 vs 進修 vs 平衡）",
            "",
            "▸ 三大支柱（英文、GPA、動機）",
            "",
            "▸ 善用獎學金（學海計畫、教育部獎學金）",
            "",
            "▸ 行前準備（簽證、保險、網卡、2張信用卡）",
            "",
            "▸ Speak Up & Work Smart",
            "",
            "▸ Try New Things & 跳脫舒適圈",
            "",
            "▸ 代表學校，為自己創造機會"
        ]
    )

    # 投影片 19: Q&A
    add_content_slide_centered(
        "Q&A 常見問題",
        [
            "Q: GPA要多少才能申請？",
            "A: 沒有絕對標準，但越高越好。學海計畫各校標準不同。",
            "",
            "Q: 英文要考到幾分？",
            "A: 多數學校：托福80+ / 雅思6.0+ / 在校英文80+",
            "",
            "Q: 沒有姊妹校可以交換嗎？",
            "A: 可以！透過教育部獎學金，不限姊妹校。",
            "",
            "Q: 交換要花多少錢？",
            "A: 視國家而定。學海計畫補助5-30萬，弱勢優秀生可全額。",
            "",
            "Q: 現在開始準備來得及嗎？",
            "A: 看你想何時出發。至少需要12個月準備時間。"
        ]
    )

    # 投影片 20: 結尾
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 背景
    background = slide.shapes.add_shape(
        1,
        Inches(0), Inches(0),
        Inches(13.333), Inches(7.5)
    )
    background.fill.solid()
    background.fill.fore_color.rgb = COLOR_PRIMARY
    background.line.fill.background()

    # 主要文字
    main_box = slide.shapes.add_textbox(
        Inches(1), Inches(2.5),
        Inches(11.333), Inches(1.5)
    )
    tf = main_box.text_frame
    tf.text = "感謝聆聽！"
    p = tf.paragraphs[0]
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.alignment = PP_ALIGN.CENTER

    # 副文字
    sub_box = slide.shapes.add_textbox(
        Inches(1), Inches(4.2),
        Inches(11.333), Inches(0.8)
    )
    tf = sub_box.text_frame
    tf.text = "祝各位交換順利，收穫滿滿！"
    p = tf.paragraphs[0]
    p.font.size = Pt(28)
    p.font.color.rgb = COLOR_SECONDARY
    p.alignment = PP_ALIGN.CENTER

    # 英文標語
    quote_box = slide.shapes.add_textbox(
        Inches(1), Inches(5.5),
        Inches(11.333), Inches(0.6)
    )
    tf = quote_box.text_frame
    tf.text = "The world is your classroom."
    p = tf.paragraphs[0]
    p.font.size = Pt(24)
    p.font.italic = True
    p.font.color.rgb = COLOR_ACCENT
    p.alignment = PP_ALIGN.CENTER

    # 儲存簡報
    output_file = '/home/user/czech-exchange/交換經驗分享會_蔡秀吉_百川學程.pptx'
    prs.save(output_file)
    print(f"✓ 簡報已成功創建：{output_file}")
    print(f"✓ 共 {len(prs.slides)} 張投影片")
    print(f"✓ 尺寸：16:9 (13.333\" x 7.5\")")
    print(f"✓ 設計風格：現代、專業、置中對齊")

    return output_file

if __name__ == "__main__":
    create_presentation()
