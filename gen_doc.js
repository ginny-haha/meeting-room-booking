const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  TabStopPosition, TableOfContents
} = require('docx');
const fs = require('fs');

const BLUE = "1F4E79";
const LIGHT_BLUE = "D6E4F0";
const MID_BLUE = "2E75B6";
const GRAY_BG = "F2F2F2";
const DARK_GRAY = "404040";
const WHITE = "FFFFFF";

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 30, bold: true, color: BLUE })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: MID_BLUE })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: DARK_GRAY })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80, line: 360 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "222222", ...opts })]
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "222222" })]
  });
}
function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "222222" })]
  });
}
function empty(spaceBefore = 80) {
  return new Paragraph({ spacing: { before: spaceBefore, after: 0 }, children: [] });
}

function headerRow(cells, widths) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((text, i) => new TableCell({
      borders: cellBorders,
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, font: "Arial", size: 20, bold: true, color: WHITE })]
      })]
    }))
  });
}
function dataRow(cells, widths, shade = false) {
  return new TableRow({
    children: cells.map((text, i) => new TableCell({
      borders: cellBorders,
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: shade ? GRAY_BG : WHITE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        children: [new TextRun({ text, font: "Arial", size: 20, color: "222222" })]
      })]
    }))
  });
}

function infoBox(label, content) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1800, 7226],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: cellBorders,
          width: { size: 1800, type: WidthType.DXA },
          shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: 20, bold: true, color: BLUE })] })]
        }),
        new TableCell({
          borders: cellBorders,
          width: { size: 7226, type: WidthType.DXA },
          shading: { fill: WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: content, font: "Arial", size: 20, color: "222222" })] })]
        })
      ]
    })]
  });
}

const today = new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Taipei' });

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u25CF", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25CB", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]
      },
      {
        reference: "numbers",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.DECIMAL, text: "%1.%2.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: MID_BLUE },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: DARK_GRAY },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 4 } },
          spacing: { after: 120 },
          children: [
            new TextRun({ text: "內部會議室預約系統｜網頁需求規格書", font: "Arial", size: 18, color: "666666" }),
            new TextRun({ text: "　　　　　　　　　　　　　　　　　　　　　　　　　　　", font: "Arial", size: 18 }),
            new TextRun({ text: "機密文件", font: "Arial", size: 18, color: MID_BLUE, bold: true }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 4 } },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          spacing: { before: 120 },
          children: [
            new TextRun({ text: "© 公司資訊部　版本 v1.0　" + today, font: "Arial", size: 16, color: "888888" }),
            new TextRun({ text: "\t第 ", font: "Arial", size: 16, color: "888888" }),
            new PageNumber({ font: "Arial", size: 16, color: "888888" }),
            new TextRun({ text: " 頁", font: "Arial", size: 16, color: "888888" }),
          ]
        })]
      })
    },
    children: [
      // ─── 封面區塊 ───
      empty(1440),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [new TextRun({ text: "內部會議室預約系統", font: "Arial", size: 64, bold: true, color: BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "網頁功能需求規格書", font: "Arial", size: 40, color: MID_BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 720 },
        children: [new TextRun({ text: "Meeting Room Booking System — Requirements Specification", font: "Arial", size: 22, color: "888888", italics: true })]
      }),
      empty(240),
      new Table({
        width: { size: 6000, type: WidthType.DXA },
        columnWidths: [2200, 3800],
        alignment: AlignmentType.CENTER,
        rows: [
          ["文件版本", "v1.0（初版）"],
          ["建立日期", today],
          ["文件狀態", "待審核"],
          ["需求提出單位", "人事暨行政部"],
          ["負責開發單位", "資訊部"],
          ["適用對象", "全體員工"],
        ].map(([l, v], i) => new TableRow({
          children: [
            new TableCell({
              borders: cellBorders,
              width: { size: 2200, type: WidthType.DXA },
              shading: { fill: i === 0 ? BLUE : LIGHT_BLUE, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: l, font: "Arial", size: 20, bold: true, color: i === 0 ? WHITE : BLUE })] })]
            }),
            new TableCell({
              borders: cellBorders,
              width: { size: 3800, type: WidthType.DXA },
              shading: { fill: WHITE, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: v, font: "Arial", size: 20, color: "222222" })] })]
            })
          ]
        }))
      }),
      empty(2880),

      // ─── 目錄 ───
      new Paragraph({
        pageBreakBefore: true,
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 0, after: 240 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE, space: 4 } },
        children: [new TextRun({ text: "目　錄", font: "Arial", size: 30, bold: true, color: BLUE })]
      }),
      new TableOfContents("目錄", { hyperlink: true, headingStyleRange: "1-3" }),
      empty(360),

      // ─── 一、背景與目的 ───
      new Paragraph({ pageBreakBefore: true, children: [] }),
      h1("一、背景與目的"),
      body("本公司近期完成新人事系統導入，原系統附屬之會議室預約功能已隨之停用。為維持跨部門會議、徵才面試、客戶拜訪等日常活動之順暢運作，避免多部門同時使用同一會議室造成衝突，特提出開發「輕量級內部會議室預約網頁」之需求。"),
      empty(),
      body("本系統預計以獨立網頁形式部署，供全體同仁透過內部網路瀏覽器存取，無需安裝任何軟體或申請額外帳號，降低使用門檻。"),
      empty(160),

      h2("1.1 問題描述"),
      bullet("新人事系統不含會議室預約功能，舊有流程已失效"),
      bullet("現階段同仁透過口頭或 LINE 群組預約，容易發生撞期"),
      bullet("缺乏統一的資訊揭露平台，各部門無法即時掌握會議室使用狀態"),
      empty(120),

      h2("1.2 預期效益"),
      bullet("全公司統一使用單一平台預約，消除撞期問題"),
      bullet("即時可視化看板，降低溝通成本"),
      bullet("支援循環預約，減少週期性會議之重複操作"),
      bullet("自助式取消流程，提升會議室資源利用率"),
      empty(240),

      // ─── 二、系統範圍 ───
      h1("二、系統範圍與基本設定"),

      h2("2.1 預約標的（會議室清單）"),
      body("本系統共管理以下五間會議室："),
      empty(80),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [600, 2500, 2500, 3426],
        rows: [
          headerRow(["#", "會議室名稱", "所在地點", "備註"], [600, 2500, 2500, 3426]),
          dataRow(["1", "台北大會議室", "台北辦公室", "大型會議使用"], [600, 2500, 2500, 3426], false),
          dataRow(["2", "台北小會議室", "台北辦公室", "小型討論、面試使用"], [600, 2500, 2500, 3426], true),
          dataRow(["3", "彰化大會議室", "彰化辦公室", "大型會議使用"], [600, 2500, 2500, 3426], false),
          dataRow(["4", "彰化小會議室", "彰化辦公室", "小型討論、面試使用"], [600, 2500, 2500, 3426], true),
          dataRow(["5", "越南AP會議室", "越南廠區（AP）", "時間顯示統一以台灣時間 GMT+8 為基準"], [600, 2500, 2500, 3426], false),
        ]
      }),
      empty(160),

      h2("2.2 時間設定"),
      bullet("預約時間單位：以「30分鐘」為一個最小預約單位"),
      bullet("每日開放預約時段：06:00 ～ 20:00（共28個時間格）"),
      bullet("時段範例：09:00–09:30、09:30–10:00、10:00–10:30……依此類推"),
      bullet("所有時間顯示均以台灣標準時間（GMT+8）為基準，包含越南AP會議室"),
      empty(240),

      // ─── 三、功能需求 ───
      h1("三、功能需求規格"),

      h2("3.1 功能一：主畫面看板（Dashboard）"),

      h3("3.1.1 日期選擇器"),
      bullet("網頁頂部設有日期導覽列，預設顯示當日日期"),
      bullet("提供「上一天」、「下一天」按鈕進行日期切換"),
      bullet("提供日期輸入欄位（date picker），允許跳轉至任意日期"),
      bullet("設有「今天」快捷按鈕，一鍵回到當日"),
      bullet("日期顯示格式：YYYY/MM/DD（星期），例如：2025/05/13（二）"),
      empty(100),

      h3("3.1.2 預約看板版面"),
      bullet("版面以一日為單位呈現"),
      bullet("橫軸（欄）：顯示5間會議室名稱，作為欄標題"),
      bullet("縱軸（列）：顯示時間，由06:00至20:00，以30分鐘為一列（共28列）"),
      bullet("整點時間列（如07:00、08:00）以較深色/粗邊框加以區分"),
      bullet("狀態顯示規則如下：", ),
      bullet("閒置格：顯示為空白可點擊格，滑鼠移過時有視覺回饋（hover效果）", 1),
      bullet("已預約格：以醒目色塊填滿，直接顯示「申請部門-預約目的」（例如：人資部-月會）", 1),
      bullet("每間會議室使用獨立色系進行區分，方便視覺識別", 1),
      bullet("循環預約之色塊另顯示「循環」標籤", 1),
      empty(240),

      h2("3.2 功能二：新增預約"),

      h3("3.2.1 觸發方式"),
      body("點擊主畫面任一閒置時段格子，系統即於畫面頂層彈出預約填寫視窗（模態視窗），無需捲動頁面。"),
      empty(100),

      h3("3.2.2 預約表單欄位"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 1600, 2200, 3226],
        rows: [
          headerRow(["欄位名稱", "是否必填", "預設值/格式", "說明"], [2000, 1600, 2200, 3226]),
          dataRow(["申請人姓名", "必填", "空白", "自由輸入文字"], [2000, 1600, 2200, 3226], false),
          dataRow(["申請人部門", "必填", "空白", "自由輸入文字（如：人資部）"], [2000, 1600, 2200, 3226], true),
          dataRow(["預約時段", "必填", "自動帶入點擊時段", "下拉選單，可延長至當日最晚可用時段"], [2000, 1600, 2200, 3226], false),
          dataRow(["預約目的", "必填", "空白", "自由輸入（如：週會、面試、客戶拜訪）"], [2000, 1600, 2200, 3226], true),
          dataRow(["取消密碼", "必填", "空白", "申請人自訂4位數字，用於日後取消驗證"], [2000, 1600, 2200, 3226], false),
          dataRow(["循環設定", "選填", "單次（預設）", "選項：單次、每週循環、每兩週循環"], [2000, 1600, 2200, 3226], true),
          dataRow(["循環結束日期", "循環時必填", "空白", "選擇循環終止日期（需晚於當日）"], [2000, 1600, 2200, 3226], false),
        ]
      }),
      empty(120),

      h3("3.2.3 表單操作按鈕"),
      bullet("【取消】：關閉視窗，不儲存任何資料，畫面維持不變"),
      bullet("【完成預約】：送出表單，執行防撞期檢查後儲存"),
      empty(100),

      h3("3.2.4 防撞期機制（並發控制）"),
      body("點擊【完成預約】時，系統須執行以下即時檢查："),
      empty(80),
      numbered("系統即時查詢資料庫，確認所選時段是否仍為閒置狀態"),
      numbered("若時段仍為閒置：儲存預約，關閉視窗，看板即時更新為已預約色塊"),
      numbered("若時段已被他人搶先預約：不儲存本次資料，跳出警告提示（範例：「很遺憾，此時段剛剛已被其他同仁預約，請選擇其他時段」），並自動重新整理看板顯示最新狀態"),
      empty(240),

      h2("3.3 功能三：循環預約"),

      h3("3.3.1 循環選項"),
      bullet("單次預約（預設選項）"),
      bullet("每週循環：從所選日期起，每隔7天重複預約相同時段"),
      bullet("每兩週循環：從所選日期起，每隔14天重複預約相同時段"),
      empty(100),

      h3("3.3.2 循環終止條件"),
      bullet("選擇循環時，必須指定「循環結束日期」"),
      bullet("結束日期不可早於或等於開始日期"),
      bullet("系統依結束日期自動計算循環日期清單後批次寫入資料庫"),
      empty(100),

      h3("3.3.3 循環防撞期處理"),
      body("系統在建立循環預約時，若偵測到未來某些循環日期已被他人預約，處理邏輯如下："),
      empty(80),
      bullet("有衝突的日期：略過不建立預約"),
      bullet("無衝突的日期：正常建立預約"),
      bullet("完成後跳出提示，說明哪些日期因衝突未能建立（範例：「10/15 及 10/22 的時段已被占用，其餘日期均已為您預約完成」）"),
      bullet("若所有日期均衝突：不建立任何預約，提示用戶另選時段"),
      empty(240),

      h2("3.4 功能四：查看與取消預約"),

      h3("3.4.1 查看預約詳情"),
      bullet("點擊看板上任一已預約色塊，彈出詳情視窗（模態視窗）"),
      bullet("詳情視窗顯示：會議室名稱、日期、時段、申請人、部門、預約目的"),
      bullet("若屬循環預約，另顯示：循環頻率、循環結束日期"),
      empty(100),

      h3("3.4.2 取消預約流程"),
      numbered("詳情視窗下方提供【取消此預約】按鈕"),
      numbered("點擊後要求輸入當初設定的4位數字取消密碼（分為4格輸入框）"),
      numbered("密碼驗證正確：刪除預約紀錄，時段恢復為閒置，看板即時更新"),
      numbered("密碼驗證錯誤：顯示錯誤提示，不執行任何刪除動作"),
      empty(100),

      h3("3.4.3 循環預約取消選項"),
      body("若所取消之預約屬循環預約，密碼驗證通過後系統進一步詢問取消範圍："),
      empty(80),
      bullet("選項A：僅取消本日（當天該筆紀錄），其餘循環日期維持不變"),
      bullet("選項B：取消本日及之後所有循環預約（本日以後的循環全部刪除）"),
      empty(240),

      // ─── 四、非功能需求 ───
      h1("四、非功能需求"),

      h2("4.1 技術規格建議"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          headerRow(["項目", "說明"], [2500, 6526]),
          dataRow(["前端技術", "HTML5 / CSS3 / JavaScript（或 React 等現代框架）"], [2500, 6526], false),
          dataRow(["後端技術", "Node.js、Python（Django/FastAPI）或 PHP 等，由資訊部自行決定"], [2500, 6526], true),
          dataRow(["資料庫", "MySQL、PostgreSQL 或 SQLite（依部署環境而定）"], [2500, 6526], false),
          dataRow(["部署環境", "公司內部伺服器（Intranet），限內網存取"], [2500, 6526], true),
          dataRow(["瀏覽器相容性", "Chrome / Edge 最新版本（主要）；Firefox / Safari 次要支援"], [2500, 6526], false),
          dataRow(["裝置支援", "電腦瀏覽器為主，建議具基本響應式設計支援平板"], [2500, 6526], true),
          dataRow(["時區處理", "所有時間儲存與顯示均以 GMT+8（台灣時間）為基準"], [2500, 6526], false),
        ]
      }),
      empty(160),

      h2("4.2 效能需求"),
      bullet("頁面初始載入時間：3秒以內（正常網路環境）"),
      bullet("預約送出至看板更新：2秒以內完成"),
      bullet("支援同時使用人數：預估最高50人同時線上"),
      empty(160),

      h2("4.3 安全性需求"),
      bullet("取消密碼採雜湊（Hash）儲存，不得以明文寫入資料庫"),
      bullet("所有輸入欄位需進行基本的 XSS 與 SQL Injection 防護"),
      bullet("系統部署限縮於內部網路，不對外開放"),
      bullet("無需使用者登入驗證（採開放式存取設計）"),
      empty(240),

      // ─── 五、UI/UX 要求 ───
      h1("五、使用者介面（UI/UX）要求"),

      h2("5.1 整體風格"),
      bullet("介面風格簡潔、專業，易於識別"),
      bullet("五間會議室使用不同色系區分，顏色於圖例中標示"),
      bullet("整點時間列以視覺方式明顯標記（如加粗邊框或背景色）"),
      bullet("已預約時段色塊內直接顯示「部門-目的」文字，超出長度截斷加省略號"),
      empty(120),

      h2("5.2 模態視窗規範"),
      bullet("所有彈出視窗（預約填寫、預約詳情）均採模態視窗形式"),
      bullet("開啟視窗時，背景加半透明遮罩（遮蔽但可辨識底層看板）"),
      bullet("視窗須出現於當前可視區域頂部，不可要求使用者捲動頁面才能看到"),
      bullet("點擊遮罩背景可關閉視窗"),
      empty(120),

      h2("5.3 提示與回饋"),
      bullet("所有操作結果（預約成功、衝突警告、密碼錯誤等）以 Toast 提示訊息呈現"),
      bullet("Toast 顯示於畫面頂端，3.5秒後自動消失"),
      bullet("不同狀態使用不同顏色（成功：綠色；錯誤：紅色；警告：橘色）"),
      empty(240),

      // ─── 六、資料結構建議 ───
      h1("六、資料結構建議"),
      body("以下為建議之資料庫表格結構，資訊部可依實際技術選型調整："),
      empty(120),

      h2("6.1 預約紀錄表（bookings）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 1800, 1500, 3726],
        rows: [
          headerRow(["欄位名稱", "資料型別", "是否必填", "說明"], [2000, 1800, 1500, 3726]),
          dataRow(["id", "VARCHAR / UUID", "是", "唯一識別碼（主鍵）"], [2000, 1800, 1500, 3726], false),
          dataRow(["room_idx", "INT", "是", "會議室編號（0～4）"], [2000, 1800, 1500, 3726], true),
          dataRow(["start_slot", "VARCHAR", "是", "開始時段（如 09:00）"], [2000, 1800, 1500, 3726], false),
          dataRow(["duration", "INT", "是", "時段數量（1單位=30分鐘）"], [2000, 1800, 1500, 3726], true),
          dataRow(["name", "VARCHAR", "是", "申請人姓名"], [2000, 1800, 1500, 3726], false),
          dataRow(["dept", "VARCHAR", "是", "申請人部門"], [2000, 1800, 1500, 3726], true),
          dataRow(["purpose", "VARCHAR", "是", "預約目的"], [2000, 1800, 1500, 3726], false),
          dataRow(["pin_hash", "VARCHAR", "是", "取消密碼（Hash後儲存）"], [2000, 1800, 1500, 3726], true),
          dataRow(["recur_type", "VARCHAR", "是", "once / weekly / biweekly"], [2000, 1800, 1500, 3726], false),
          dataRow(["recur_end", "DATE", "否", "循環結束日期（單次預約為null）"], [2000, 1800, 1500, 3726], true),
          dataRow(["created_at", "DATETIME", "是", "建立時間"], [2000, 1800, 1500, 3726], false),
        ]
      }),
      empty(160),

      h2("6.2 預約日期表（booking_dates）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 1800, 1500, 3726],
        rows: [
          headerRow(["欄位名稱", "資料型別", "是否必填", "說明"], [2000, 1800, 1500, 3726]),
          dataRow(["id", "INT", "是", "自動遞增主鍵"], [2000, 1800, 1500, 3726], false),
          dataRow(["booking_id", "VARCHAR", "是", "關聯至 bookings.id（外鍵）"], [2000, 1800, 1500, 3726], true),
          dataRow(["date", "DATE", "是", "預約日期（YYYY-MM-DD）"], [2000, 1800, 1500, 3726], false),
          dataRow(["is_cancelled", "BOOLEAN", "是", "是否已取消（預設 false）"], [2000, 1800, 1500, 3726], true),
        ]
      }),
      empty(240),

      // ─── 七、驗收標準 ───
      h1("七、驗收標準"),

      h2("7.1 功能驗收清單"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 4526, 1500, 2500],
        rows: [
          headerRow(["#", "驗收項目", "優先級", "驗收方式"], [500, 4526, 1500, 2500]),
          dataRow(["1", "日期切換功能正常，可前後切換及跳轉", "高", "手動測試"], [500, 4526, 1500, 2500], false),
          dataRow(["2", "看板正確顯示5間會議室與28個時間格", "高", "手動測試"], [500, 4526, 1500, 2500], true),
          dataRow(["3", "閒置格可點擊，彈出預約視窗於可視區域頂部", "高", "手動測試"], [500, 4526, 1500, 2500], false),
          dataRow(["4", "所有必填欄位驗證正常，未填時不允許送出", "高", "手動測試"], [500, 4526, 1500, 2500], true),
          dataRow(["5", "預約成功後看板即時更新", "高", "手動測試"], [500, 4526, 1500, 2500], false),
          dataRow(["6", "同時點擊相同時段，後者收到衝突提示", "高", "並發測試"], [500, 4526, 1500, 2500], true),
          dataRow(["7", "循環預約正確產生所有循環日期", "高", "資料驗證"], [500, 4526, 1500, 2500], false),
          dataRow(["8", "循環預約部分衝突時，正確略過並提示", "中", "手動測試"], [500, 4526, 1500, 2500], true),
          dataRow(["9", "點擊已預約色塊正確顯示詳情", "高", "手動測試"], [500, 4526, 1500, 2500], false),
          dataRow(["10", "密碼正確可成功取消，錯誤時提示錯誤", "高", "手動測試"], [500, 4526, 1500, 2500], true),
          dataRow(["11", "循環取消選項（本日/全部）運作正確", "中", "手動測試"], [500, 4526, 1500, 2500], false),
          dataRow(["12", "越南AP會議室時間顯示與其他房間一致（GMT+8）", "中", "手動驗證"], [500, 4526, 1500, 2500], true),
          dataRow(["13", "取消密碼以 Hash 形式儲存於資料庫", "高", "資料庫檢查"], [500, 4526, 1500, 2500], false),
          dataRow(["14", "頁面在 Chrome / Edge 最新版正常顯示", "高", "跨瀏覽器測試"], [500, 4526, 1500, 2500], true),
        ]
      }),
      empty(240),

      // ─── 八、專案時程建議 ───
      h1("八、專案時程建議"),
      body("以下時程為參考建議，實際執行時程請由資訊部依人力與資源評估後確認："),
      empty(120),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 2500, 2000, 4026],
        rows: [
          headerRow(["階段", "工作項目", "建議工期", "產出物"], [500, 2500, 2000, 4026]),
          dataRow(["1", "需求確認與技術選型", "3 個工作天", "確認後需求書、技術架構文件"], [500, 2500, 2000, 4026], false),
          dataRow(["2", "資料庫設計與後端開發", "5 個工作天", "資料庫 Schema、API 規格"], [500, 2500, 2000, 4026], true),
          dataRow(["3", "前端介面開發", "5 個工作天", "可運作前端頁面"], [500, 2500, 2000, 4026], false),
          dataRow(["4", "整合測試與錯誤修正", "3 個工作天", "測試報告"], [500, 2500, 2000, 4026], true),
          dataRow(["5", "使用者驗收測試（UAT）", "2 個工作天", "驗收簽核文件"], [500, 2500, 2000, 4026], false),
          dataRow(["6", "正式上線與教育訓練", "1 個工作天", "上線通知、操作說明"], [500, 2500, 2000, 4026], true),
          dataRow(["合計", "—", "約 19 個工作天", "—"], [500, 2500, 2000, 4026], false),
        ]
      }),
      empty(240),

      // ─── 九、其他說明 ───
      h1("九、其他說明與限制"),

      h2("9.1 不在本次需求範圍內"),
      bullet("使用者登入 / 帳號管理功能（採開放式存取）"),
      bullet("電子郵件通知或行事曆整合（如 Google Calendar、Outlook）"),
      bullet("行動裝置 App（以網頁瀏覽器為主要載體）"),
      bullet("報表匯出或統計分析功能"),
      bullet("歷史預約紀錄查詢功能"),
      empty(120),

      h2("9.2 未來可擴充功能（供參考）"),
      bullet("整合公司 AD/LDAP 進行身分驗證"),
      bullet("預約確認 Email 通知功能"),
      bullet("行動裝置響應式介面優化"),
      bullet("會議室使用率統計報表"),
      bullet("管理員後台（強制取消、批次管理）"),
      empty(160),

      h2("9.3 需求窗口聯絡資訊"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 6526],
        rows: [
          headerRow(["項目", "說明"], [2500, 6526]),
          dataRow(["需求提出部門", "人事暨行政部"], [2500, 6526], false),
          dataRow(["資訊部承辦窗口", "請資訊部指定對應人員"], [2500, 6526], true),
          dataRow(["文件版本管理", "本文件如有修訂，請更新版本號並通知相關人員"], [2500, 6526], false),
          dataRow(["問題反應管道", "請透過公司內部工單系統提出"), [2500, 6526], true),
        ]
      }),
      empty(360),

      // ─── 簽核欄 ───
      h1("十、文件簽核"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2256, 2256, 2257, 2257],
        rows: [
          headerRow(["角色", "姓名", "部門", "簽核日期"], [2256, 2256, 2257, 2257]),
          dataRow(["需求提出人", "", "人事暨行政部", ""], [2256, 2256, 2257, 2257], false),
          dataRow(["部門主管審核", "", "人事暨行政部", ""], [2256, 2256, 2257, 2257], true),
          dataRow(["資訊部承辦", "", "資訊部", ""], [2256, 2256, 2257, 2257], false),
          dataRow(["資訊部主管", "", "資訊部", ""], [2256, 2256, 2257, 2257], true),
        ]
      }),
      empty(200),
      body("本文件經上述人員簽核後正式生效，如需修訂請重新走簽核流程。", { color: "888888", italics: true }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/會議室預約系統_網頁需求書.docx', buf);
  console.log('done');
});
