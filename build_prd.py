from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from datetime import date

out='PRD_個人行程與任務管理系統.docx'
doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.72); sec.bottom_margin=Inches(.68); sec.left_margin=Inches(.78); sec.right_margin=Inches(.78)
styles=doc.styles
for nm,size,color in [('Normal',9.5,'303447'),('Title',27,'202333'),('Heading 1',17,'202333'),('Heading 2',12,'5142BA'),('Heading 3',10,'303447')]:
    st=styles[nm]; st.font.name='Microsoft JhengHei'; st.font.size=Pt(size); st.font.color.rgb=RGBColor.from_string(color)
    st._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
    if nm.startswith('Heading'): st.font.bold=True
styles['Normal'].paragraph_format.space_after=Pt(5)
styles['Normal'].paragraph_format.line_spacing=1.12
styles['Title'].font.bold=True
styles['Title'].paragraph_format.space_after=Pt(6)
styles['Heading 1'].paragraph_format.space_before=Pt(15); styles['Heading 1'].paragraph_format.space_after=Pt(6)
styles['Heading 2'].paragraph_format.space_before=Pt(9); styles['Heading 2'].paragraph_format.space_after=Pt(4)
# Header/footer
header=sec.header.paragraphs[0]; header.text='產品需求文件  /  個人行程與任務管理系統'; header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
for r in header.runs: r.font.name='Microsoft JhengHei'; r.font.size=Pt(8); r.font.color.rgb=RGBColor.from_string('777B8D'); r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=footer.add_run('日序 Daymark  ·  需求討論版  ·  2026-10-30'); r.font.size=Pt(8);r.font.color.rgb=RGBColor.from_string('85899A');r.font.name='Microsoft JhengHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')

def shade(cell,fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd');shd.set(qn('w:fill'),fill);tcPr.append(shd)
def set_cell_border(cell,color='D9DCE5',sz='5'):
    tc=cell._tc;tcPr=tc.get_or_add_tcPr();borders=tcPr.first_child_found_in('w:tcBorders')
    if borders is None: borders=OxmlElement('w:tcBorders');tcPr.append(borders)
    for edge in ('top','left','bottom','right','insideH','insideV'):
        el=OxmlElement('w:'+edge);el.set(qn('w:val'),'single');el.set(qn('w:sz'),sz);el.set(qn('w:color'),color);borders.append(el)
def set_repeat_table_header(row):
    trPr=row._tr.get_or_add_trPr();el=OxmlElement('w:tblHeader');el.set(qn('w:val'),'true');trPr.append(el)
def table(headers,rows,widths=None):
    t=doc.add_table(rows=1, cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    if widths:
        for col,w in zip(t.columns,widths): col.width=Inches(w)
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i];c.text=h;shade(c,'28304A');c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for p in c.paragraphs:
            p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(4)
            for r in p.runs:r.font.name='Microsoft JhengHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');r.font.size=Pt(8.5);r.font.bold=True;r.font.color.rgb=RGBColor(255,255,255)
        set_cell_border(c)
    set_repeat_table_header(t.rows[0])
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        for i,val in enumerate(row):
            cells[i].text=str(val);cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths: cells[i].width=Inches(widths[i])
            if ri%2:shade(cells[i],'F5F6FA')
            set_cell_border(cells[i])
            for p in cells[i].paragraphs:
                p.paragraph_format.space_before=Pt(3);p.paragraph_format.space_after=Pt(3);p.paragraph_format.line_spacing=1.08
                for r in p.runs:r.font.name='Microsoft JhengHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');r.font.size=Pt(8.2);r.font.color.rgb=RGBColor.from_string('303447')
    doc.add_paragraph().paragraph_format.space_after=Pt(1)
    return t
def para(text='',style=None,boldlead=None):
    p=doc.add_paragraph(style=style)
    if boldlead and text.startswith(boldlead):
        p.add_run(boldlead).bold=True;p.add_run(text[len(boldlead):])
    else:p.add_run(text)
    return p
def bullets(items):
    for x in items:
        p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(x)

doc.add_paragraph('產品需求文件',style='Title')
p=doc.add_paragraph();p.paragraph_format.space_after=Pt(13)
r=p.add_run('個人行程與任務管理系統');r.bold=True;r.font.size=Pt(14);r.font.color.rgb=RGBColor.from_string('5142BA')
para('文件版本：0.1（初步需求基線）    ｜    日期：2026 年 10 月 30 日    ｜    產品代稱：日序 Daymark')
para('文件目的：將目前訪談中已確認的目標、流程和規則整理成可供產品、設計與工程團隊討論的需求基線。多人共用、重複任務與執行紀錄等尚未釐清的內容會明確標為待決策事項，不視為已承諾功能。')

doc.add_heading('1. 產品概述',level=1)
para('本產品協助個人管理短、中、長期計畫與日常任務，並連結計畫目標、每日安排、任務實際完成情況及週／月進度回顧。產品需降低待辦遺漏與未完成事項失去脈絡的風險，並讓使用者知道各計畫目前推進到哪裡。未來可能擴充給家人、同事或其他人共用。')
para('核心價值：把「想完成的計畫」拆成可執行任務；保留未完成任務的完成率與備註；提供可調整的延續／重新安排流程；以週、月報告追蹤計畫任務完成率。')

doc.add_heading('2. 目標與成功指標',level=1)
table(['目標','初步衡量方式'],[
('減少任務遺漏','使用者能在每日清單與待重新安排清單中找到所有尚未完成的任務。'),
('保留執行脈絡','任務移入待重新安排時，完成率、備註、計畫關聯、截止日與原安排資訊仍可查看。'),
('看見計畫進度','可查看單一計畫目前任務完成率，並能按週／月查看各計畫任務完成狀況。'),
('降低每日排程阻力','未完成任務可依使用者為任務設定的方式延續，或移入重新安排清單。')],[2.0,4.8])
para('正式 KPI 數值與量測期間尚待產品負責人確認；上述為可用性與功能完整性的初步驗收方向。')

doc.add_heading('3. 使用者與使用情境',level=1)
table(['使用者／角色','需要完成的事'],[
('個人使用者（首要）','建立不同週期的計畫；每天安排任務；記錄任務完成狀況、完成率與備註；整理延期任務；回顧週／月進度。'),
('共享空間成員（未來）','與家人、同事或其他人共用計畫與任務；具體角色、權限、指派及通知規則待釐清。')],[1.65,5.15])
para('代表情境：使用者建立一項中期計畫，將今日任務連結至該計畫；任務進行一部分後更新完成率並留下備註。若當日未完成且有截止日，任務連同備註進入待重新安排清單，並依截止日排序；週末可檢視此計畫本週任務完成率。')

doc.add_heading('4. 範圍',level=1)
doc.add_heading('4.1 初版範圍（已確認）',level=2)
bullets(['計畫管理：新增、檢視、編輯、刪除計畫；計畫由使用者手動標記短期、中期或長期，不以固定天數判定。','任務管理：新增、檢視、編輯、刪除任務；任務可連結一項計畫，也可不屬於任何計畫。','每日任務：依安排日期查看任務，可瀏覽前後日期，並區分待完成／已完成。','任務執行紀錄：任務具完成狀態、完成率與備註。標記已完成時完成率自動為 100%；未完成任務可設定 0–99% 完成率。','未完成處理：任務可設定隔天自動排入或移至待重新安排。設有截止日的未完成任務移至待重新安排清單。','待重新安排：保留任務與備註，按截止日期遠近排序；使用者可選擇新日期後重新排入每日清單。','計畫進度：顯示計畫所屬任務的進度摘要。','週／月報告：切換週或月、瀏覽前期／本期，顯示整體任務數與完成率，以及各計畫任務數、已完成數、完成率與平均進度。','本機原型資料保存於瀏覽器；正式產品的帳號、雲端持久化、跨裝置同步與備份尚未定義。'])
doc.add_heading('4.2 暫不納入／待確認',level=2)
bullets(['多人共享與權限（已有未來需求方向，但成員邀請、指派、編輯權限、共用範圍與通知未確認）。','重複任務：使用者表示重複任務不用每次輸入；重複週期、例外日期、生成時點、單次完成與系列修改規則待確認。','完整日曆／外部行事曆同步、提醒通知、附件、標籤、搜尋與匯出。','實際開始與結束時間、時間追蹤、每日／週期歷史快照和任務狀態稽核紀錄。','多位執行者、子任務、任務依賴、評論與活動紀錄。'])

doc.add_heading('5. 資訊架構與主要畫面',level=1)
table(['畫面／模組','主要內容與動作'],[
('今日總覽','每日任務清單、日期切換、待完成／已完成篩選、任務新增、今日統計、每週節奏摘要、進行中計畫摘要、待重新安排摘要。'),
('計畫管理','計畫清單與週期標籤篩選；建立／編輯計畫；查看計畫說明、目標日期、任務數與目前進度；由計畫新增關聯任務。'),
('待重新安排','未完成任務清單；截止日近者優先；展示完成率、備註及所屬計畫；選擇新日期後重新安排。'),
('週月報告','週／月切換；期間前後導覽與回到本期；整體指標卡；各計畫完成率比較表。'),
('設定','說明未完成處理、完成率計算、計畫進度邏輯；未來共用空間規劃資訊。')],[1.5,5.3])

doc.add_heading('6. 功能需求與驗收條件',level=1)
table(['ID','需求','驗收條件'],[
('FR-01','管理計畫','使用者可建立／編輯計畫名稱、短／中／長期標籤、說明與選填目標日期；刪除計畫時，原任務保留並解除計畫關聯。'),
('FR-02','管理任務','可輸入任務名稱、安排日期、選填時間、所屬計畫（可空白）、選填截止日、完成率、備註及未完成處理方式。'),
('FR-03','完成狀態與完成率','勾選完成後完成率自動成為 100%；取消完成後回復先前未完成率或 0%。未完成率限制為 0–99%。'),
('FR-04','任務備註','新增或更新備註後，任務進入待重新安排時仍完整保留。'),
('FR-05','未完成自動延續','使用者選隔天自動排入且無截止日，任務安排日期順延一天，保留進度／備註／計畫關聯。'),
('FR-06','待重新安排','使用者選移至待重新安排，任務未完成時進入該清單；有截止日的未完成任務亦進入該清單，優先順序依截止日升冪。無截止日項目排在有截止日項目之後。'),
('FR-07','重新安排任務','使用者選定新日期並確認後，任務回到每日任務；任務內容、完成率、備註、計畫關聯保留，完成狀態維持未完成。'),
('FR-08','計畫進度','計畫頁與總覽顯示所屬任務數及進度；任務完成率依未完成任務百分比與已完成 100% 計算。具體聚合方式需和報告指標定義一致。'),
('FR-09','週月報告','使用者可切換每週／每月、前後期間與本期；系統依任務安排日期篩選期間內任務，呈現整體數據與逐計畫明細。'),
('FR-10','刪除任務','使用者可刪除任務；刪除前提示確認。已刪除任務不納入報告與計畫進度。')],[.65,1.45,4.7])

doc.add_heading('7. 業務規則與指標定義',level=1)
table(['項目','初步定義'],[
('計畫週期','短期／中期／長期為使用者自行選擇的分類標籤；不由開始日、截止日或固定時間長度推算。'),
('完成狀態','布林狀態；已完成任務的完成率固定為 100%。取消完成時依介面／產品決策恢復先前進度，或回到未完成進度。'),
('未完成進度','0–99%；可用於記錄部分完成程度，並保留於延期任務。'),
('計畫目前進度（原型建議）','計畫所屬有效任務的完成率平均值；若無任務則為 0%。此值代表任務平均進度，不等同「完成任務比例」。'),
('週／月計畫完成率','已完成任務數 ÷ 本期該計畫安排任務總數 × 100%；本期無任務時顯示 0% 或「—」，正式呈現待確認。'),
('週／月平均進度','本期該計畫任務完成率的算術平均；無任務時顯示 0% 或「—」，正式呈現待確認。'),
('報告日期範圍','每週以週一至週日；每月以曆月；依任務安排日期歸期。需支援的時區及週起始日待確認。'),
('未完成任務規則','有截止日的未完成任務進待重新安排，不論該任務原設定的延續方式；沒有截止日則依任務自選方式處理。'),
('報告內容','整體摘要可包含無計畫任務；逐計畫列只計入該計畫的任務，無計畫任務不歸入任何計畫。')],[1.65,5.15])
para('重要口徑提醒：介面需區分「完成率」（已完成數／任務數）與「平均進度」（任務完成率平均）。兩者不可混用，避免部分完成任務讓報表完成率被高估。')

doc.add_heading('8. 資料需求（初步）',level=1)
table(['實體','主要欄位'],[
('計畫 Plan','plan_id、名稱、週期標籤（short／medium／long）、說明、目標日期（選填）、建立／更新時間、狀態。'),
('任務 Task','task_id、名稱、安排日期、時間（選填）、截止日期（選填）、計畫 ID（選填）、完成率、完成狀態、備註、未完成處理方式、任務狀態（每日／待重新安排／已刪除）、建立／更新時間。'),
('使用者／工作區（正式版待定）','user_id、workspace_id、顯示名稱、成員角色、工作區關聯；須於多人共享規格確認後細化。'),
('報告（計算結果）','原則上可依任務資料即時計算；如需歷史快照或稽核報告，須另定保存策略與重算規則。')],[1.7,5.1])

doc.add_heading('9. 非功能需求',level=1)
bullets(['可用性：新增任務應能快速完成；重要欄位分組清楚，未完成處理規則可理解。','響應式：桌面與手機均可操作每日清單、編輯任務、查看待重新安排及週月報告。','資料正確性：完成狀態、完成率與報告口徑一致；變更任務後計畫進度與週月報告即時更新。','資料保存：正式版需確認登入、雲端保存、跨裝置同步、備份與刪除策略；目前瀏覽器原型僅保存本機資料。','可存取性：互動元件需有可辨識名稱、鍵盤操作、足夠色彩對比及非色彩的狀態提示。','隱私與權限：未來共享後，須定義個人資料隔離、共享工作區權限與資料匯出／刪除流程。'])

doc.add_heading('10. 主要流程',level=1)
doc.add_heading('10.1 建立計畫與任務',level=2)
para('計畫管理 → 新增計畫 → 輸入名稱與短／中／長期分類（可填說明與目標日期）→ 儲存 → 由計畫新增任務或於任務表單選擇所屬計畫 → 設定日期、進度、備註、截止日與未完成處理方式 → 儲存。任務也可選「不屬於計畫」。')
doc.add_heading('10.2 未完成任務處理',level=2)
para('每日任務未完成時，若有截止日，轉入待重新安排並按截止日排序；沒有截止日則按任務個別設定順延至隔天或轉入待重新安排。重新安排時選擇新的安排日期，原任務的備註與進度保留。已完成任務不參與未完成延續。')
doc.add_heading('10.3 查看週／月報告',level=2)
para('由左側導覽進入週月報告 → 選每週或每月 → 瀏覽目前或前後期 → 查看整體任務摘要與各計畫明細。期間按任務安排日期分組；每筆任務屬於最多一個計畫或未歸屬任何計畫。')

doc.add_heading('11. 邊界情況與錯誤處理',level=1)
bullets(['截止日早於安排日期：建議允許儲存但明確提示日期衝突，或阻止儲存；需確認產品決策。','任務被改為已完成：從待重新安排清單移除，完成率設為 100%，仍保留任務紀錄。','任務日期改到未來：是否仍留在待重新安排，或因重新排期而自動恢復每日任務，需明確規定。','刪除計畫：不連帶刪除任務；解除任務計畫關聯並保留任務。','重新安排任務日期落在已完成週／月：歷史報告依目前任務日期計算可能改變；是否需要固定歷史快照待確認。','同一任務需不需要跨日分拆或重複生成：目前資料模型以單一任務記錄處理，重複任務系列規則待設計。'])

doc.add_heading('12. 依賴與風險',level=1)
table(['項目','影響／處理建議'],[
('自動延續的執行時機','瀏覽器關閉時不能單靠前端自動觸發；正式版需由伺服器排程或於使用者下次登入時計算並補處理。'),
('週月報告口徑','完成率與平均進度容易混淆；需在 UI 明確標示分子分母及計算期間。'),
('完成任務重新開啟','恢復完成前進度或重設為 0% 會影響報告，需明確制定。'),
('共享功能擴充','單人資料模型需保留 workspace／owner 擴充點，但不應在未確認權限規格前假設共用流程。'),
('歷史資料與報告','若任務日期／狀態可修改，過往報告會動態變更；若報告需作為正式歷史紀錄，需追加快照機制。')],[1.8,5.0])

doc.add_heading('13. 待決策問題',level=1)
table(['優先','待確認事項','建議方向'],[
('P0','週定義是否為週一至週日？時區是否以使用者設定為準？','預設週一至週日，所有日期以工作區時區計算。'),
('P0','計畫進度要呈現任務平均進度、已完成任務比例，還是兩者？','同時呈現：計畫總進度＝平均進度；另列完成任務數／任務總數。'),
('P0','重新安排後任務是否退出待重新安排並恢復每日執行狀態？','是；依使用者選定的新日期回到每日任務。'),
('P1','未完成且已逾截止日的任務是否仍依截止日升冪排列？','建議逾期任務置頂，先依最早截止日，再依建立時間排序。'),
('P1','重複任務週期及跳過／停止系列規則？','先確定每日、每週、每月等頻率及例外處理，再納入規格。'),
('P1','共享工作區成員角色、指派、通知與個人／共享任務界線？','先定義個人工作區與共享工作區，任務需有負責人與建立者欄位。'),
('P2','是否需要實際開始／結束時間、每日歷史快照、報告匯出？','核心任務流程穩定後再評估。')],[.55,3.15,3.1])

doc.add_heading('14. 原型現況與限制',level=1)
para('目前已有單頁 HTML/CSS/JavaScript 原型，可操作每日任務、計畫管理、待重新安排與週／月報告，並以瀏覽器 localStorage 保存示範資料。此原型不包含帳號登入、伺服器自動排程、多人同步、通知或資料備份。日期自動轉期目前以頁面開啟時檢查舊日期任務作為概念展示，不代表正式排程服務。')

doc.add_heading('附錄 A. 需求追溯',level=1)
table(['需求來源（訪談）','對應需求'],[
('未來可能與家人、同事或其他人共用','角色與共享工作區作為後續擴充；權限待定。'),
('計畫不受固定時間長短限制，手動標為短／中／長期','FR-01、業務規則：計畫週期。'),
('任務可自動隔天排入或移至待重新安排','FR-05、FR-06。'),
('有截止日未完成任務移入重新安排，依截止日由近到遠','FR-06、待重新安排畫面。'),
('已完成、刪除任務、完成率；完成即 100%','FR-03、FR-10。'),
('任務有備註，未完成時任務與備註一起重新安排','FR-04、FR-06。'),
('任務可歸屬計畫，也可不屬於計畫；顯示計畫進度','FR-02、FR-08。'),
('新增每週、每月報告統計各計畫任務完成率','FR-09、週／月報告畫面及指標定義。')],[3.2,3.6])

# Keep rows intact where possible, and normalize paragraph font for CJK
for p in doc.paragraphs:
    for r in p.runs:
        r.font.name='Microsoft JhengHei'; r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
doc.core_properties.title='個人行程與任務管理系統 產品需求文件'
doc.core_properties.subject='產品需求規格 PRD 初步需求基線'
doc.core_properties.author='產品需求訪談整理'
doc.save(out)
print(out)
