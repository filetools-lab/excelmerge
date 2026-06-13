const fileInput = document.querySelector('#file-input');
const dropZone = document.querySelector('.drop-zone');
const fileList = document.querySelector('#file-list');
const mergeButton = document.querySelector('#merge-button');
const resetButton = document.querySelector('#reset-button');
const exportButton = document.querySelector('#export-button');
const statusText = document.querySelector('#status');
const statsTableBody = document.querySelector('#stats-table tbody');
const previewTable = document.querySelector('#preview-table');
const statFiles = document.querySelector('#stat-files');
const statRows = document.querySelector('#stat-rows');
const statColumns = document.querySelector('#stat-columns');
const statNumeric = document.querySelector('#stat-numeric');
const statProcessTime = document.querySelector('#stat-process-time');
const statTimeSaved = document.querySelector('#stat-time-saved');
const statSheets = document.querySelector('#stat-sheets');
const languageSelect = document.querySelector('#language-select');
const pendingCount = document.querySelector('#pending-count');
const validationModal = document.querySelector('#validation-modal');
const validationList = document.querySelector('#validation-list');
const reviewFilesButton = document.querySelector('#review-files-button');
const continueMergeButton = document.querySelector('#continue-merge-button');
const shareButton = document.querySelector('#share-button');
const demoButton = document.querySelector('#demo-button');
const analyticsVisits = document.querySelector('#analytics-visits');
const analyticsMerges = document.querySelector('#analytics-merges');
const analyticsFiles = document.querySelector('#analytics-files');
const analyticsExports = document.querySelector('#analytics-exports');

let selectedFiles = [];
let sourceSheets = [];
let mergedRows = [];
let summaryRows = [];
let allHeaders = [];
let currentLanguage = 'en';
let pendingValidationSheets = [];
let pendingValidationDurationMs = 0;
let mergeMetrics = { durationMs: 0, timeSavedMinutes: 0 };

const SOURCE_FILE = 'Source File';
const SOURCE_SHEET = 'Source Sheet';
const SITE_URL = 'https://filetools-lab.github.io/excelmerge/';
const ANALYTICS_KEY = 'filetoolsLabExcelMergeStats';
const translations = {
  en: {
    pageTitle: 'Merge Multiple Excel Files in Seconds - Local Excel Merge Tool',
    brand: 'FileTools Lab',
    brandTagline: 'Excel merge tools for private reports',
    languageLabel: 'Language',
    shareButton: 'Copy link',
    shareCopied: 'Copied',
    privacyBadge: 'Files stay local',
    eyebrow: 'Browser based Excel merge tool',
    heroTitle: 'Merge Multiple Excel Files in Seconds',
    heroPointLocal: '100% Local Processing',
    heroPointUpload: 'No Upload Required',
    heroPointFree: 'Free Browser-Based Tool',
    heroCopy: 'Combine Excel and CSV reports directly in your browser. No registration. No server upload. Your files never leave your device.',
    heroStartButton: 'Start merging',
    demoDataButton: 'Try Demo Data',
    demoLoadedStatus: 'Demo data loaded. Three monthly sales reports are ready to merge.',
    privacyEyebrow: 'Privacy-first processing',
    privacyTitle: 'Your Excel files never leave this browser',
    privacyCopy: 'Files are read, checked, merged, and exported locally on your device. This tool does not upload, store, or analyze your workbook contents on a server.',
    privacyPointA: 'No upload server',
    privacyPointACopy: 'Workbook data is processed in the browser.',
    privacyPointB: 'No saved files',
    privacyPointBCopy: 'Closing the page clears the selected files.',
    privacyPointC: 'Export locally',
    privacyPointCCopy: 'The final workbook is generated on your device.',
    useCasesEyebrow: 'Perfect For',
    useCasesTitle: 'Combine multiple Excel reports for everyday business work',
    useCasesCopy: 'Use this browser based Excel merge tool for recurring files that need to become one clean workbook.',
    useCaseDaily: 'Daily Reports',
    useCaseWeekly: 'Weekly Reports',
    useCaseMonthly: 'Monthly Reports',
    useCaseSales: 'Sales Reports',
    useCaseStore: 'Store Performance Reports',
    useCaseFinancial: 'Financial Worksheets',
    stepsEyebrow: 'How it works',
    stepsTitle: 'Merge Excel files in four steps',
    stepsCopy: 'Add files, check differences, choose an output format, and export a clean workbook.',
    stepOneTitle: 'Add files',
    stepOneCopy: 'Add .xlsx, .xls, or .csv files in one or more batches.',
    stepTwoTitle: 'Review checks',
    stepTwoCopy: 'See warnings for mismatched headers, empty sheets, or content-type differences.',
    stepThreeTitle: 'Choose merge type',
    stepThreeCopy: 'Merge all rows into one sheet or keep source sheets in one workbook.',
    stepFourTitle: 'Export workbook',
    stepFourCopy: 'Download summary.xlsx with a cover page, stats, and merged data.',
    demoTitle: 'See the workflow before you upload',
    demoCopy: 'Three separate tables with five rows each become one merged table with fifteen rows.',
    beforeLabel: 'Before',
    afterLabel: 'After',
    demoFileA: 'North.xlsx',
    demoFileB: 'South.xlsx',
    demoFileC: 'Online.csv',
    demoFileAMeta: '2 sheets · 1,248 rows',
    demoFileBMeta: '2 sheets · 1,180 rows',
    demoFileCMeta: '1 sheet · 508 rows',
    beforeSummary: '3 tables · 5 rows each',
    demoArrow: 'Merge',
    demoTableA: 'Table A',
    demoTableB: 'Table B',
    demoTableC: 'Table C',
    afterSummary: '1 table · 15 rows',
    afterDetail: 'Rows are appended into one sheet with Source File and Source Sheet columns.',
    demoFileAShort: 'Sales_Jan, Sales_Feb',
    demoFileBShort: 'Retail, Wholesale',
    demoFileCShort: 'Online Orders',
    sheetSalesJan: 'Sales_Jan',
    sheetSalesJanRows: '624 rows',
    sheetSalesFeb: 'Sales_Feb',
    sheetSalesFebRows: '624 rows',
    sheetRetail: 'Retail',
    sheetRetailRows: '760 rows',
    sheetWholesale: 'Wholesale',
    sheetWholesaleRows: '420 rows',
    sheetOnlineOrders: 'Online Orders',
    sheetOnlineOrdersRows: '508 rows',
    outputWorkbook: 'summary.xlsx',
    outputWorkbookMeta: '5 source sheets · 2,936 merged rows',
    coverPageTitle: 'Cover Page summary',
    coverTab: 'Cover Page',
    statsTab: 'Summary Stats',
    sheetLink: 'linked summary',
    sheetStatsDetail: 'column metrics',
    dataTab: 'Merged Data',
    uploadTitle: '1. Upload reports',
    uploadCopy: 'Select one or more Excel files, or try demo data first. Every worksheet in each workbook is merged locally.',
    workspaceEyebrow: 'Merge Workspace',
    workspaceTitle: 'Add files and choose how to merge them',
    workspaceCopy: 'This is the only area where you select files, choose the merge type, and start processing.',
    dropTitle: 'Choose Excel files',
    dropHint: 'or drag and drop them here. Files are added locally and are not uploaded.',
    pendingLabel: 'Pending files',
    pendingHint: 'Add files in batches. Merging starts only after you click Merge files.',
    mergeType: 'Merge type',
    oneSheet: 'One sheet',
    oneSheetCopy: 'Combine every row into a single worksheet.',
    workbookSheets: 'Workbook sheets',
    workbookSheetsCopy: 'Keep each source worksheet as a separate workbook sheet.',
    mergeButton: 'Merge files',
    resetButton: 'Reset',
    waitingStatus: 'Waiting for files.',
    readyStatus: (count) => `${count} file${count === 1 ? '' : 's'} ready to merge.`,
    addedStatus: (added, total) => `Added ${added} file${added === 1 ? '' : 's'}. ${total} file${total === 1 ? '' : 's'} pending merge.`,
    duplicateStatus: (total) => `Those files are already in the pending list. ${total} file${total === 1 ? '' : 's'} pending merge.`,
    readingStatus: 'Reading workbooks locally in your browser...',
    sheetJsError: 'SheetJS could not be loaded. Check your internet connection or vendor xlsx.full.min.js locally.',
    mergedStatus: (rows, files, duration, saved) => `Merged ${rows.toLocaleString()} row${rows === 1 ? '' : 's'} from ${files} file${files === 1 ? '' : 's'}. Completed in ${duration}. Estimated ${saved} saved.`,
    noRowsStatus: 'No data rows were found in the selected workbooks.',
    mergeError: (message) => `Unable to merge files: ${message}`,
    statFiles: 'Files Processed',
    statRows: 'Rows Merged',
    statColumns: 'Columns',
    statNumeric: 'Numeric columns',
    statProcessTime: 'Process time',
    statTimeSaved: 'Estimated Time Saved',
    statSheetsCombined: 'Sheets combined',
    resultsEyebrow: 'Merge Results',
    resultsTitle: 'Review the output after merging',
    resultsCopy: 'These cards and tables update only after you run a merge or load demo data.',
    statsTitle: '2. Summary statistics',
    statsCopy: 'Numeric columns include count, sum, average, minimum, and maximum values.',
    exportButton: 'Export summary.xlsx',
    columnHeader: 'Column',
    typeHeader: 'Type',
    nonEmptyHeader: 'Non-empty',
    blankHeader: 'Blank',
    uniqueHeader: 'Unique',
    sumHeader: 'Sum',
    averageHeader: 'Average',
    minHeader: 'Min',
    maxHeader: 'Max',
    statsEmpty: 'Merge files to calculate statistics.',
    previewTitle: '3. Merged data preview',
    previewCopy: 'Previewing the first 100 merged rows. The exported workbook contains all rows.',
    previewEmpty: 'No merged rows yet.',
    faqEyebrow: 'FAQ',
    faqTitle: 'Common questions',
    faqCopy: 'Answers about merging Excel files locally, offline use, file privacy, and browser limits.',
    faqXlsxQuestion: 'Can I merge XLSX files?',
    faqXlsxAnswer: 'Yes. This free browser based Excel merge tool supports XLSX, XLS and CSV files.',
    faqUploadQuestion: 'Are files uploaded?',
    faqUploadAnswer: 'No. All processing happens locally in your browser, so you can merge Excel files without uploading sensitive data.',
    faqOfflineQuestion: 'Does this work offline?',
    faqOfflineAnswer: 'Yes. After loading the page you can continue working offline in the same browser tab.',
    faqSecureQuestion: 'Is my data secure?',
    faqSecureAnswer: 'Yes. Your files never leave your computer, which makes this useful for financial worksheets, sales reports and private business data.',
    faqLimitQuestion: 'Is there a file size limit?',
    faqLimitAnswer: 'There is no server file size limit. The practical limit is your browser memory and device performance.',
    professionalEyebrow: 'Professional Features',
    professionalTitle: 'Coming Soon',
    professionalCopy: 'All current Excel merge features remain free. Future professional tools may help teams analyze reports and create recurring business outputs faster.',
    proAi: 'AI Report Analysis',
    proWeekly: 'Weekly Report Generation',
    proMonthly: 'Monthly Report Generation',
    proPpt: 'PowerPoint Export',
    proTemplates: 'Advanced Templates',
    proBatch: 'Batch Processing',
    blogEyebrow: 'Excel Guides',
    blogTitle: 'Learn how to merge Excel reports safely',
    blogCopy: 'Read practical guides about local Excel merging, sales report workflows, and when to use Power Query.',
    blogArticleOne: 'How to Merge Multiple Excel Files Without VBA',
    blogArticleTwo: 'How to Combine Daily Sales Reports into One Excel File',
    blogArticleThree: 'Excel Merge vs Power Query',
    blogArticleFour: 'Merge Excel Files Locally Without Uploading Data',
    analyticsEyebrow: 'Usage Snapshot',
    analyticsTitle: 'Local activity in this browser',
    analyticsCopy: 'These privacy-friendly counters are stored only in this browser. Connect Google Analytics, Plausible, or Umami later for real public traffic numbers.',
    analyticsLocalLabel: 'Local browser stats',
    analyticsVisits: 'Visits',
    analyticsMerges: 'Merge Operations',
    analyticsFiles: 'Files Processed',
    analyticsExports: 'Exports Generated',
    footerBrand: 'Powered by FileTools Lab',
    footerRoadmap: 'Excel Merge · CSV Merge · Excel Compare · CSV Compare · Excel To CSV · CSV To Excel · AI Report Analyzer',
    numericType: 'Numeric',
    textType: 'Text/Mixed',
    coverTitle: 'Local Excel Report Merger',
    coverMergeType: 'Merge type',
    coverOneSheet: 'One sheet',
    coverWorkbookSheets: 'Workbook sheets',
    coverFilesMerged: 'Files merged',
    coverSourceSheets: 'Source worksheets merged',
    coverTotalRows: 'Total merged rows',
    coverProcessingTime: 'Processing time',
    coverEstimatedTimeSaved: 'Estimated time saved',
    coverGeneratedAt: 'Generated at',
    coverOutputSheet: 'Output sheet',
    coverSourceFile: 'Source file',
    coverSourceSheet: 'Source sheet',
    coverRows: 'Rows',
    validationEyebrow: 'Review needed',
    validationTitle: 'Some tables do not match',
    validationCopy: 'The files can still be merged, but these differences may affect the output. Please review them first.',
    validationHeaderIssue: 'Header mismatch',
    validationTypeIssue: 'Content type mismatch',
    validationEmptyIssue: 'Empty or unreadable sheet',
    validationMore: (count) => `And ${count} more issue${count === 1 ? '' : 's'} not shown.`,
    reviewFilesButton: 'Return to check',
    continueMergeButton: 'Continue merge',
    validationReviewStatus: 'Merge paused. Review the pending files, then run merge again.',
    validationContinueStatus: 'Continuing merge with the listed differences...',
    missingColumns: (columns) => `Missing columns: ${columns}`,
    extraColumns: (columns) => `Extra columns: ${columns}`,
    orderMismatch: 'Column order differs from the first sheet.',
    typeMismatch: (column, expected, actual) => `"${column}" looks like ${actual}, but the first sheet looks like ${expected}.`,
    emptySheetDetail: 'No data rows were found in this sheet.',
    issueLocation: (file, sheet) => `${file} / ${sheet}`,
    typeNumeric: 'numeric data',
    typeDate: 'date data',
    typeText: 'text data',
    typeBlank: 'blank data',
  },
  zh: {
    pageTitle: '快速合并多个 Excel 文件 - 本地 Excel 合并工具',
    brand: 'FileTools Lab',
    brandTagline: '面向隐私报表的 Excel 合并工具',
    languageLabel: '语言',
    shareButton: '复制链接',
    shareCopied: '已复制',
    privacyBadge: '文件仅保留在本地',
    eyebrow: '基于浏览器的 Excel 合并工具',
    heroTitle: '几秒内合并多个 Excel 文件',
    heroPointLocal: '100% 本地处理',
    heroPointUpload: '无需上传文件',
    heroPointFree: '免费的浏览器工具',
    heroCopy: '直接在浏览器中合并 Excel 和 CSV 报表。无需注册，无需上传到服务器，文件不会离开你的设备。',
    heroStartButton: '开始合并',
    demoDataButton: '试用演示数据',
    demoLoadedStatus: '演示数据已载入。3 个月度销售报表已准备好合并。',
    privacyEyebrow: '隐私优先',
    privacyTitle: '你的 Excel 文件不会离开当前浏览器',
    privacyCopy: '文件读取、检查、合并和导出都在你的设备本地完成。这个工具不会把工作簿内容上传到服务器，也不会存储或分析你的表格数据。',
    privacyPointA: '没有上传服务器',
    privacyPointACopy: '工作簿数据直接在浏览器里处理。',
    privacyPointB: '不保存文件',
    privacyPointBCopy: '关闭页面后，已选择的文件会被清空。',
    privacyPointC: '本地导出',
    privacyPointCCopy: '最终工作簿在你的设备上生成。',
    useCasesEyebrow: '适合场景',
    useCasesTitle: '为日常业务工作合并多个 Excel 报表',
    useCasesCopy: '适合将周期性报表合并成一个清晰工作簿，例如日报、周报、月报和销售报表。',
    useCaseDaily: '日报',
    useCaseWeekly: '周报',
    useCaseMonthly: '月报',
    useCaseSales: '销售报表',
    useCaseStore: '门店经营报表',
    useCaseFinancial: '财务工作表',
    stepsEyebrow: '使用步骤',
    stepsTitle: '四步完成 Excel 合并',
    stepsCopy: '添加文件、检查差异、选择合并方式，然后导出清晰的汇总工作簿。',
    stepOneTitle: '添加文件',
    stepOneCopy: '分批添加 .xlsx、.xls 或 .csv 文件。',
    stepTwoTitle: '检查差异',
    stepTwoCopy: '合并前查看表头不一致、空 sheet 或内容类型差异。',
    stepThreeTitle: '选择合并方式',
    stepThreeCopy: '把所有行合并到一个 sheet，或把来源 sheet 保留在一个 workbook 中。',
    stepFourTitle: '导出工作簿',
    stepFourCopy: '下载包含封面页、统计页和合并数据的 summary.xlsx。',
    demoTitle: '上传前先看工作效果',
    demoCopy: '3 张各 5 行的数据表，会被合并成 1 张 15 行的汇总表。',
    beforeLabel: '合并前',
    afterLabel: '合并后',
    demoFileA: '北区.xlsx',
    demoFileB: '南区.xlsx',
    demoFileC: '线上.csv',
    demoFileAMeta: '2 个 sheet · 1,248 行',
    demoFileBMeta: '2 个 sheet · 1,180 行',
    demoFileCMeta: '1 个 sheet · 508 行',
    beforeSummary: '3 张表 · 每张 5 行',
    demoArrow: '合并',
    demoTableA: '表 A',
    demoTableB: '表 B',
    demoTableC: '表 C',
    afterSummary: '1 张表 · 15 行',
    afterDetail: '所有行向下追加到一个 sheet，并保留 Source File 和 Source Sheet 列。',
    demoFileAShort: '一月销售、二月销售',
    demoFileBShort: '零售、批发',
    demoFileCShort: '线上订单',
    sheetSalesJan: '一月销售',
    sheetSalesJanRows: '624 行',
    sheetSalesFeb: '二月销售',
    sheetSalesFebRows: '624 行',
    sheetRetail: '零售',
    sheetRetailRows: '760 行',
    sheetWholesale: '批发',
    sheetWholesaleRows: '420 行',
    sheetOnlineOrders: '线上订单',
    sheetOnlineOrdersRows: '508 行',
    outputWorkbook: 'summary.xlsx',
    outputWorkbookMeta: '5 个来源 sheet · 2,936 行合并数据',
    coverPageTitle: 'Cover Page 汇总',
    coverTab: 'Cover Page',
    statsTab: 'Summary Stats',
    sheetLink: '带链接的汇总',
    sheetStatsDetail: '列统计指标',
    dataTab: '合并数据',
    uploadTitle: '1. 上传报表',
    uploadCopy: '选择一个或多个 Excel 文件，也可以先试用演示数据。每个工作簿里的所有 worksheet 都会在本地读取。',
    workspaceEyebrow: '合并工作台',
    workspaceTitle: '添加文件并选择合并方式',
    workspaceCopy: '这里是唯一的操作区域：选择文件、选择合并类型，然后开始处理。',
    dropTitle: '选择 Excel 文件',
    dropHint: '也可以把文件拖拽到这里。文件只会加入本地队列，不会上传。',
    pendingLabel: '待合并文件',
    pendingHint: '可以分批添加文件。只有点击“开始合并”后才会运行。',
    mergeType: '合并类型',
    oneSheet: '合并到一个 sheet',
    oneSheetCopy: '把所有行合并到同一个 worksheet 中。',
    workbookSheets: '合并到多个 sheet',
    workbookSheetsCopy: '把每个来源 worksheet 保留为导出工作簿里的独立 sheet。',
    mergeButton: '开始合并',
    resetButton: '重置',
    waitingStatus: '等待选择文件。',
    readyStatus: (count) => `已选择 ${count} 个文件，可以开始合并。`,
    addedStatus: (added, total) => `已添加 ${added} 个文件，当前共有 ${total} 个文件待合并。`,
    duplicateStatus: (total) => `这些文件已经在待合并列表中，当前共有 ${total} 个文件待合并。`,
    readingStatus: '正在浏览器本地读取工作簿...',
    sheetJsError: 'SheetJS 加载失败。请检查网络连接，或将 xlsx.full.min.js 改成本地文件。',
    mergedStatus: (rows, files, duration, saved) => `已从 ${files} 个文件中合并 ${rows.toLocaleString()} 行数据。处理用时 ${duration}，预计节省 ${saved}。`,
    noRowsStatus: '所选工作簿中没有找到数据行。',
    mergeError: (message) => `无法合并文件：${message}`,
    statFiles: '已处理文件',
    statRows: '已合并行数',
    statColumns: '列数',
    statNumeric: '数字列',
    statProcessTime: '处理耗时',
    statTimeSaved: '预计节省时间',
    statSheetsCombined: '合并 Sheet 数',
    resultsEyebrow: '合并结果',
    resultsTitle: '合并后在这里查看输出结果',
    resultsCopy: '这些卡片和表格只会在你运行合并或载入演示数据后更新。',
    statsTitle: '2. 汇总统计',
    statsCopy: '数字列会统计数量、总和、平均值、最小值和最大值。',
    exportButton: '导出 summary.xlsx',
    columnHeader: '列名',
    typeHeader: '类型',
    nonEmptyHeader: '非空',
    blankHeader: '空白',
    uniqueHeader: '唯一值',
    sumHeader: '总和',
    averageHeader: '平均值',
    minHeader: '最小值',
    maxHeader: '最大值',
    statsEmpty: '合并文件后会在这里生成统计信息。',
    previewTitle: '3. 合并数据预览',
    previewCopy: '这里预览前 100 行。导出的工作簿会包含全部行。',
    previewEmpty: '还没有合并数据。',
    faqEyebrow: '常见问题',
    faqTitle: '常见问题',
    faqCopy: '关于本地合并 Excel、离线使用、文件隐私和浏览器限制的说明。',
    faqXlsxQuestion: '可以合并 XLSX 文件吗？',
    faqXlsxAnswer: '可以。这个免费的浏览器 Excel 合并工具支持 XLSX、XLS 和 CSV 文件。',
    faqUploadQuestion: '文件会被上传吗？',
    faqUploadAnswer: '不会。所有处理都在你的浏览器本地完成，你可以在不上传敏感数据的情况下合并 Excel 文件。',
    faqOfflineQuestion: '这个工具可以离线使用吗？',
    faqOfflineAnswer: '可以。页面加载完成后，你可以在同一个浏览器标签页里继续离线处理。',
    faqSecureQuestion: '我的数据安全吗？',
    faqSecureAnswer: '是的。文件不会离开你的电脑，因此适合处理财务工作表、销售报表和私密业务数据。',
    faqLimitQuestion: '有文件大小限制吗？',
    faqLimitAnswer: '没有服务器文件大小限制，实际限制取决于你的浏览器内存和设备性能。',
    professionalEyebrow: '专业功能',
    professionalTitle: '即将推出',
    professionalCopy: '当前 Excel 合并功能保持免费。未来专业工具会帮助团队分析报表，并更快生成周期性业务输出。',
    proAi: 'AI 报表分析',
    proWeekly: '周报生成',
    proMonthly: '月报生成',
    proPpt: 'PowerPoint 导出',
    proTemplates: '高级模板',
    proBatch: '批量处理',
    blogEyebrow: 'Excel 指南',
    blogTitle: '学习如何安全合并 Excel 报表',
    blogCopy: '阅读关于本地 Excel 合并、销售报表流程，以及何时使用 Power Query 的实用指南。',
    blogArticleOne: '如何不使用 VBA 合并多个 Excel 文件',
    blogArticleTwo: '如何将每日销售报表合并成一个 Excel 文件',
    blogArticleThree: 'Excel Merge 与 Power Query 对比',
    blogArticleFour: '无需上传数据，本地合并 Excel 文件',
    analyticsEyebrow: '使用统计',
    analyticsTitle: '此浏览器中的本地活动',
    analyticsCopy: '这些隐私友好的计数只保存在当前浏览器中。以后接入 Google Analytics、Plausible 或 Umami 后，再展示真实的公开访问数据。',
    analyticsLocalLabel: '本浏览器统计',
    analyticsVisits: '访问次数',
    analyticsMerges: '合并次数',
    analyticsFiles: '处理文件数',
    analyticsExports: '导出次数',
    footerBrand: 'Powered by FileTools Lab',
    footerRoadmap: 'Excel Merge · CSV Merge · Excel Compare · CSV Compare · Excel To CSV · CSV To Excel · AI Report Analyzer',
    numericType: '数字',
    textType: '文本/混合',
    coverTitle: '本地 Excel 报表合并工具',
    coverMergeType: '合并类型',
    coverOneSheet: '合并到一个 sheet',
    coverWorkbookSheets: '合并到多个 sheet',
    coverFilesMerged: '合并文件数',
    coverSourceSheets: 'Sheet 数',
    coverTotalRows: '合并总行数',
    coverProcessingTime: '处理耗时',
    coverEstimatedTimeSaved: '预计节省时间',
    coverGeneratedAt: '生成时间',
    coverOutputSheet: '导出 sheet',
    coverSourceFile: '来源文件',
    coverSourceSheet: '来源 sheet',
    coverRows: '行数',
    validationEyebrow: '需要确认',
    validationTitle: '发现表格不一致',
    validationCopy: '这些文件仍然可以合并，但差异可能影响导出结果。请先确认是否继续。',
    validationHeaderIssue: '表头不一致',
    validationTypeIssue: '内容类型不一致',
    validationEmptyIssue: '空表或无法读取',
    validationMore: (count) => `还有 ${count} 个问题未展示。`,
    reviewFilesButton: '返回检查',
    continueMergeButton: '继续合并',
    validationReviewStatus: '已暂停合并。请检查待合并文件，然后再次开始合并。',
    validationContinueStatus: '将带着上述差异继续合并...',
    missingColumns: (columns) => `缺少列：${columns}`,
    extraColumns: (columns) => `新增列：${columns}`,
    orderMismatch: '列顺序与第一个 sheet 不一致。',
    typeMismatch: (column, expected, actual) => `“${column}”看起来是${actual}，但第一个 sheet 中是${expected}。`,
    emptySheetDetail: '这个 sheet 中没有读取到数据行。',
    issueLocation: (file, sheet) => `${file} / ${sheet}`,
    typeNumeric: '数字',
    typeDate: '日期',
    typeText: '文本',
    typeBlank: '空白',
  },
};

const languageAttributes = {
  en: 'en',
  zh: 'zh-CN',
  hi: 'hi',
  ja: 'ja',
  pt: 'pt-BR',
  id: 'id',
};

Object.assign(translations, {
  hi: {
    ...translations.en,
    pageTitle: 'Excel मर्ज टूल - Excel और CSV फाइलें स्थानीय रूप से मर्ज करें',
    brand: 'Excel रिपोर्ट मर्जर',
    brandTagline: 'स्थानीय वर्कबुक टूल',
    languageLabel: 'भाषा',
    shareButton: 'लिंक कॉपी करें',
    shareCopied: 'कॉपी हो गया',
    privacyBadge: 'फाइलें स्थानीय रहती हैं',
    eyebrow: 'ब्राउज़र में चलने वाला Excel टूल',
    heroTitle: 'Excel रिपोर्ट को एक सारांश वर्कबुक में मर्ज करें',
    heroCopy: 'कई <strong>.xlsx</strong>, <strong>.xls</strong>, या <strong>.csv</strong> फाइलें अपलोड करें, संयुक्त पंक्तियां देखें, सारांश आंकड़े जांचें, और <strong>summary.xlsx</strong> निर्यात करें। सारी प्रोसेसिंग आपके ब्राउज़र में स्थानीय रूप से होती है।',
    privacyEyebrow: 'प्राइवेसी-फर्स्ट प्रोसेसिंग',
    privacyTitle: 'आपकी Excel फाइलें इस ब्राउज़र से बाहर नहीं जातीं',
    privacyCopy: 'फाइलें आपके डिवाइस पर ही पढ़ी, जांची, मर्ज और निर्यात की जाती हैं। यह टूल आपकी वर्कबुक सामग्री को सर्वर पर अपलोड, स्टोर या विश्लेषण नहीं करता।',
    privacyPointA: 'कोई अपलोड सर्वर नहीं',
    privacyPointACopy: 'वर्कबुक डेटा ब्राउज़र में प्रोसेस होता है।',
    privacyPointB: 'फाइलें सेव नहीं होतीं',
    privacyPointBCopy: 'पेज बंद करने पर चुनी गई फाइलें साफ हो जाती हैं।',
    privacyPointC: 'स्थानीय निर्यात',
    privacyPointCCopy: 'अंतिम वर्कबुक आपके डिवाइस पर बनती है।',
    stepsEyebrow: 'कैसे काम करता है',
    stepsTitle: 'चार चरणों में Excel फाइलें मर्ज करें',
    stepsCopy: 'फाइलें जोड़ें, अंतर जांचें, आउटपुट प्रकार चुनें, और साफ वर्कबुक निर्यात करें।',
    stepOneTitle: 'फाइलें जोड़ें',
    stepOneCopy: '.xlsx, .xls या .csv फाइलें एक या अधिक बैच में जोड़ें।',
    stepTwoTitle: 'जांच देखें',
    stepTwoCopy: 'हेडर mismatch, खाली sheet या content-type अंतर की चेतावनी देखें।',
    stepThreeTitle: 'मर्ज प्रकार चुनें',
    stepThreeCopy: 'सभी rows एक sheet में मर्ज करें या source sheets को एक workbook में रखें।',
    stepFourTitle: 'वर्कबुक निर्यात करें',
    stepFourCopy: 'Cover page, stats और merged data वाली summary.xlsx डाउनलोड करें।',
    uploadTitle: '1. रिपोर्ट अपलोड करें',
    uploadCopy: 'एक या अधिक Excel फाइलें चुनें। हर workbook की सभी worksheets पढ़ी जाएंगी।',
    dropTitle: 'Excel फाइलें चुनें',
    dropHint: 'या उन्हें यहां ड्रैग और ड्रॉप करें। फाइलें स्थानीय रूप से जोड़ी जाती हैं, अपलोड नहीं होतीं।',
    pendingLabel: 'लंबित फाइलें',
    pendingHint: 'फाइलें बैच में जोड़ें। Merge files क्लिक करने के बाद ही मर्ज शुरू होगा।',
    mergeType: 'मर्ज प्रकार',
    oneSheet: 'एक sheet',
    oneSheetCopy: 'हर row को एक worksheet में मिलाएं।',
    workbookSheets: 'Workbook sheets',
    workbookSheetsCopy: 'हर source worksheet को अलग workbook sheet के रूप में रखें।',
    mergeButton: 'फाइलें मर्ज करें',
    resetButton: 'रीसेट',
    waitingStatus: 'फाइलों की प्रतीक्षा है।',
    readyStatus: (count) => `${count} फाइल मर्ज के लिए तैयार।`,
    addedStatus: (added, total) => `${added} फाइल जोड़ी गई। कुल ${total} फाइलें मर्ज के लिए लंबित।`,
    duplicateStatus: (total) => `ये फाइलें पहले से सूची में हैं। कुल ${total} फाइलें लंबित।`,
    readingStatus: 'वर्कबुक आपके ब्राउज़र में स्थानीय रूप से पढ़ी जा रही हैं...',
    mergedStatus: (rows, files, duration, saved) => `${files} फाइलों से ${rows.toLocaleString()} rows मर्ज हुईं। समय ${duration}। अनुमानित बचत ${saved}।`,
    noRowsStatus: 'चुनी गई workbooks में data rows नहीं मिलीं।',
    mergeError: (message) => `फाइलें मर्ज नहीं हो सकीं: ${message}`,
    statFiles: 'फाइलें',
    statRows: 'Merged rows',
    statColumns: 'Columns',
    statNumeric: 'Numeric columns',
    statProcessTime: 'Process time',
    statTimeSaved: 'Time saved est.',
    statsTitle: '2. सारांश आंकड़े',
    statsCopy: 'Numeric columns में count, sum, average, minimum और maximum शामिल हैं।',
    exportButton: 'summary.xlsx निर्यात करें',
    columnHeader: 'Column',
    typeHeader: 'Type',
    nonEmptyHeader: 'Non-empty',
    blankHeader: 'Blank',
    uniqueHeader: 'Unique',
    sumHeader: 'Sum',
    averageHeader: 'Average',
    minHeader: 'Min',
    maxHeader: 'Max',
    statsEmpty: 'आंकड़े निकालने के लिए फाइलें मर्ज करें।',
    previewTitle: '3. Merged data preview',
    previewCopy: 'पहली 100 merged rows दिखाई जा रही हैं। निर्यात की गई workbook में सभी rows होंगी।',
    previewEmpty: 'अभी merged rows नहीं हैं।',
    faqEyebrow: 'FAQ',
    faqTitle: 'सामान्य प्रश्न',
    faqCopy: 'Privacy, file formats, merge behavior और mismatch checks के जवाब।',
    faqUploadQuestion: 'क्या मेरी Excel फाइलें server पर upload होती हैं?',
    faqUploadAnswer: 'नहीं। फाइलें आपके browser में स्थानीय रूप से पढ़ी और मर्ज की जाती हैं।',
    faqFormatsQuestion: 'कौन से file formats supported हैं?',
    faqFormatsAnswer: 'आप .xlsx, .xls और .csv files मर्ज कर सकते हैं।',
    faqSheetsQuestion: 'क्या मैं multiple worksheets मर्ज कर सकता हूं?',
    faqSheetsAnswer: 'हां। सभी rows को एक worksheet में मिलाएं या source worksheets को एक workbook में अलग sheets की तरह रखें।',
    faqMismatchQuestion: 'अगर headers अलग हों तो क्या होगा?',
    faqMismatchAnswer: 'मर्ज से पहले tool file और sheet location के साथ warning दिखाता है। आप जांचने के लिए लौट सकते हैं या जारी रख सकते हैं।',
    numericType: 'Numeric',
    textType: 'Text/Mixed',
    coverTitle: 'स्थानीय Excel रिपोर्ट मर्जर',
    coverMergeType: 'मर्ज प्रकार',
    coverOneSheet: 'एक sheet',
    coverWorkbookSheets: 'Workbook sheets',
    coverFilesMerged: 'मर्ज की गई फाइलें',
    coverSourceSheets: 'Source worksheets',
    coverTotalRows: 'कुल merged rows',
    coverProcessingTime: 'Processing time',
    coverEstimatedTimeSaved: 'Estimated time saved',
    coverGeneratedAt: 'Generated at',
    coverOutputSheet: 'Output sheet',
    coverSourceFile: 'Source file',
    coverSourceSheet: 'Source sheet',
    coverRows: 'Rows',
    validationEyebrow: 'Review needed',
    validationTitle: 'कुछ tables match नहीं करतीं',
    validationCopy: 'फाइलें फिर भी मर्ज हो सकती हैं, लेकिन ये अंतर output को प्रभावित कर सकते हैं। कृपया पहले review करें।',
    validationHeaderIssue: 'Header mismatch',
    validationTypeIssue: 'Content type mismatch',
    validationEmptyIssue: 'Empty or unreadable sheet',
    validationMore: (count) => `और ${count} issue नहीं दिखाए गए।`,
    reviewFilesButton: 'जांचने लौटें',
    continueMergeButton: 'मर्ज जारी रखें',
    validationReviewStatus: 'मर्ज रुका हुआ है। Pending files जांचें, फिर दुबारा merge करें।',
    validationContinueStatus: 'दिखाए गए अंतर के साथ merge जारी है...',
    missingColumns: (columns) => `Missing columns: ${columns}`,
    extraColumns: (columns) => `Extra columns: ${columns}`,
    orderMismatch: 'Column order पहली sheet से अलग है।',
    typeMismatch: (column, expected, actual) => `"${column}" ${actual} जैसा है, लेकिन पहली sheet ${expected} जैसी है।`,
    emptySheetDetail: 'इस sheet में कोई data rows नहीं मिलीं।',
    typeNumeric: 'numeric data',
    typeDate: 'date data',
    typeText: 'text data',
    typeBlank: 'blank data',
  },
  ja: {
    ...translations.en,
    pageTitle: 'Excel 結合ツール - Excel と CSV をローカルで結合',
    brand: 'Excel レポート結合ツール',
    brandTagline: 'ローカル workbook ユーティリティ',
    languageLabel: '言語',
    shareButton: 'リンクをコピー',
    shareCopied: 'コピーしました',
    privacyBadge: 'ファイルはローカルに保持',
    eyebrow: 'ブラウザだけで使える Excel ツール',
    heroTitle: 'Excel レポートを 1 つのサマリー workbook に結合',
    heroCopy: '複数の <strong>.xlsx</strong>、<strong>.xls</strong>、<strong>.csv</strong> ファイルを追加し、結合結果と統計を確認して <strong>summary.xlsx</strong> を出力できます。処理はブラウザ内でローカルに実行されます。',
    privacyEyebrow: 'プライバシー重視',
    privacyTitle: 'Excel ファイルはこのブラウザから外に出ません',
    privacyCopy: 'ファイルの読み取り、チェック、結合、出力はすべて端末上で行われます。workbook の内容はサーバーにアップロード、保存、分析されません。',
    privacyPointA: 'アップロードサーバーなし',
    privacyPointACopy: 'workbook データはブラウザ内で処理されます。',
    privacyPointB: 'ファイルを保存しません',
    privacyPointBCopy: 'ページを閉じると選択ファイルはクリアされます。',
    privacyPointC: 'ローカル出力',
    privacyPointCCopy: '最終 workbook は端末上で生成されます。',
    stepsEyebrow: '使い方',
    stepsTitle: '4 ステップで Excel を結合',
    stepsCopy: 'ファイル追加、差分確認、出力形式選択、workbook 出力を行います。',
    stepOneTitle: 'ファイルを追加',
    stepOneCopy: '.xlsx、.xls、.csv を複数回に分けて追加できます。',
    stepTwoTitle: 'チェックを確認',
    stepTwoCopy: 'ヘッダー不一致、空 sheet、内容タイプ差異を確認します。',
    stepThreeTitle: '結合タイプを選択',
    stepThreeCopy: 'すべての行を 1 sheet にまとめるか、source sheets を 1 workbook に保持します。',
    stepFourTitle: 'workbook を出力',
    stepFourCopy: 'cover page、stats、merged data を含む summary.xlsx をダウンロードします。',
    uploadTitle: '1. レポートをアップロード',
    uploadCopy: '1 つ以上の Excel ファイルを選択します。各 workbook のすべての worksheets が読み込まれます。',
    dropTitle: 'Excel ファイルを選択',
    dropHint: 'またはここにドラッグ＆ドロップ。ファイルはローカルに追加され、アップロードされません。',
    pendingLabel: '結合待ちファイル',
    pendingHint: 'ファイルは分割して追加できます。Merge files をクリックすると結合が始まります。',
    mergeType: '結合タイプ',
    oneSheet: '1 つの sheet',
    oneSheetCopy: 'すべての行を 1 つの worksheet に結合します。',
    workbookSheets: 'Workbook sheets',
    workbookSheetsCopy: '各 source worksheet を別々の workbook sheet として保持します。',
    mergeButton: 'ファイルを結合',
    resetButton: 'リセット',
    waitingStatus: 'ファイルを待機中です。',
    readyStatus: (count) => `${count} 個のファイルを結合できます。`,
    addedStatus: (added, total) => `${added} 個のファイルを追加しました。合計 ${total} 個が結合待ちです。`,
    duplicateStatus: (total) => `これらのファイルはすでに追加されています。合計 ${total} 個が結合待ちです。`,
    readingStatus: 'ブラウザ内で workbook を読み込んでいます...',
    mergedStatus: (rows, files, duration, saved) => `${files} 個のファイルから ${rows.toLocaleString()} 行を結合しました。処理時間 ${duration}、推定節約時間 ${saved}。`,
    noRowsStatus: '選択した workbooks にデータ行がありません。',
    mergeError: (message) => `ファイルを結合できません: ${message}`,
    statFiles: 'ファイル',
    statRows: '結合行数',
    statColumns: '列数',
    statNumeric: '数値列',
    statProcessTime: '処理時間',
    statTimeSaved: '推定節約時間',
    statsTitle: '2. サマリー統計',
    statsCopy: '数値列には件数、合計、平均、最小、最大が含まれます。',
    exportButton: 'summary.xlsx を出力',
    columnHeader: '列',
    typeHeader: 'タイプ',
    nonEmptyHeader: '非空',
    blankHeader: '空白',
    uniqueHeader: 'ユニーク',
    sumHeader: '合計',
    averageHeader: '平均',
    minHeader: '最小',
    maxHeader: '最大',
    statsEmpty: 'ファイルを結合すると統計が表示されます。',
    previewTitle: '3. 結合データプレビュー',
    previewCopy: '最初の 100 行を表示します。出力 workbook にはすべての行が含まれます。',
    previewEmpty: 'まだ結合データはありません。',
    faqEyebrow: 'FAQ',
    faqTitle: 'よくある質問',
    faqCopy: 'プライバシー、形式、結合方法、差分チェックについて。',
    faqUploadQuestion: 'Excel ファイルはサーバーにアップロードされますか？',
    faqUploadAnswer: 'いいえ。ファイルはブラウザ内でローカルに読み込まれ、結合されます。',
    faqFormatsQuestion: '対応形式は？',
    faqFormatsAnswer: '.xlsx、.xls、.csv を結合できます。',
    faqSheetsQuestion: '複数 worksheets を結合できますか？',
    faqSheetsAnswer: 'はい。すべての行を 1 worksheet にまとめるか、各 source worksheet を 1 workbook 内の別 sheet として保持できます。',
    faqMismatchQuestion: 'ヘッダーが違う場合は？',
    faqMismatchAnswer: '結合前に file と sheet の場所を含む警告が表示されます。確認に戻るか、そのまま続行できます。',
    numericType: '数値',
    textType: 'テキスト/混在',
    coverTitle: 'ローカル Excel レポート結合ツール',
    coverMergeType: '結合タイプ',
    coverOneSheet: '1 つの sheet',
    coverWorkbookSheets: 'Workbook sheets',
    coverFilesMerged: '結合ファイル数',
    coverSourceSheets: 'Source worksheets',
    coverTotalRows: '結合行数',
    coverProcessingTime: '処理時間',
    coverEstimatedTimeSaved: '推定節約時間',
    coverGeneratedAt: '生成日時',
    coverOutputSheet: 'Output sheet',
    coverSourceFile: 'Source file',
    coverSourceSheet: 'Source sheet',
    coverRows: 'Rows',
    validationEyebrow: '確認が必要',
    validationTitle: '一致しない table があります',
    validationCopy: 'ファイルは結合できますが、差異が出力に影響する可能性があります。先に確認してください。',
    validationHeaderIssue: 'ヘッダー不一致',
    validationTypeIssue: '内容タイプ不一致',
    validationEmptyIssue: '空または読み取り不可の sheet',
    validationMore: (count) => `さらに ${count} 件の issue は未表示です。`,
    reviewFilesButton: '確認に戻る',
    continueMergeButton: '結合を続行',
    validationReviewStatus: '結合を一時停止しました。ファイルを確認して再実行してください。',
    validationContinueStatus: '表示された差異を含めて結合を続行しています...',
    missingColumns: (columns) => `不足列: ${columns}`,
    extraColumns: (columns) => `追加列: ${columns}`,
    orderMismatch: '列順が最初の sheet と異なります。',
    typeMismatch: (column, expected, actual) => `"${column}" は ${actual} に見えますが、最初の sheet は ${expected} です。`,
    emptySheetDetail: 'この sheet にデータ行がありません。',
    typeNumeric: '数値データ',
    typeDate: '日付データ',
    typeText: 'テキストデータ',
    typeBlank: '空白データ',
  },
  pt: {
    ...translations.en,
    pageTitle: 'Ferramenta para juntar Excel - Mescle Excel e CSV localmente',
    brand: 'Mesclador de Relatórios Excel',
    brandTagline: 'Utilitário local de workbook',
    languageLabel: 'Idioma',
    shareButton: 'Copiar link',
    shareCopied: 'Copiado',
    privacyBadge: 'Arquivos ficam locais',
    eyebrow: 'Ferramenta Excel no navegador',
    heroTitle: 'Mescle relatórios Excel em um workbook de resumo',
    heroCopy: 'Adicione vários arquivos <strong>.xlsx</strong>, <strong>.xls</strong> ou <strong>.csv</strong>, visualize as linhas combinadas, revise estatísticas e exporte <strong>summary.xlsx</strong>. O processamento acontece localmente no navegador.',
    privacyEyebrow: 'Privacidade em primeiro lugar',
    privacyTitle: 'Seus arquivos Excel não saem deste navegador',
    privacyCopy: 'Os arquivos são lidos, verificados, mesclados e exportados localmente no seu dispositivo. O conteúdo do workbook não é enviado, salvo ou analisado em servidor.',
    privacyPointA: 'Sem servidor de upload',
    privacyPointACopy: 'Os dados do workbook são processados no navegador.',
    privacyPointB: 'Sem arquivos salvos',
    privacyPointBCopy: 'Ao fechar a página, os arquivos selecionados são limpos.',
    privacyPointC: 'Exportação local',
    privacyPointCCopy: 'O workbook final é gerado no seu dispositivo.',
    stepsEyebrow: 'Como funciona',
    stepsTitle: 'Mescle arquivos Excel em quatro etapas',
    stepsCopy: 'Adicione arquivos, confira diferenças, escolha o formato e exporte um workbook limpo.',
    stepOneTitle: 'Adicionar arquivos',
    stepOneCopy: 'Adicione .xlsx, .xls ou .csv em um ou mais lotes.',
    stepTwoTitle: 'Revisar alertas',
    stepTwoCopy: 'Veja avisos de cabeçalhos diferentes, sheets vazias ou diferenças de tipo.',
    stepThreeTitle: 'Escolher tipo',
    stepThreeCopy: 'Mescle todas as linhas em uma sheet ou mantenha as source sheets em um workbook.',
    stepFourTitle: 'Exportar workbook',
    stepFourCopy: 'Baixe summary.xlsx com cover page, stats e merged data.',
    uploadTitle: '1. Enviar relatórios',
    uploadCopy: 'Selecione um ou mais arquivos Excel. Todas as worksheets de cada workbook serão lidas.',
    dropTitle: 'Escolher arquivos Excel',
    dropHint: 'ou arraste e solte aqui. Os arquivos são adicionados localmente e não são enviados.',
    pendingLabel: 'Arquivos pendentes',
    pendingHint: 'Adicione arquivos em lotes. A mesclagem começa apenas ao clicar em Merge files.',
    mergeType: 'Tipo de mesclagem',
    oneSheet: 'Uma sheet',
    oneSheetCopy: 'Combine todas as linhas em uma única worksheet.',
    workbookSheets: 'Workbook sheets',
    workbookSheetsCopy: 'Mantenha cada source worksheet como uma sheet separada.',
    mergeButton: 'Mesclar arquivos',
    resetButton: 'Redefinir',
    waitingStatus: 'Aguardando arquivos.',
    readyStatus: (count) => `${count} arquivo${count === 1 ? '' : 's'} pronto${count === 1 ? '' : 's'} para mesclar.`,
    addedStatus: (added, total) => `${added} arquivo${added === 1 ? '' : 's'} adicionado${added === 1 ? '' : 's'}. ${total} arquivo${total === 1 ? '' : 's'} pendente${total === 1 ? '' : 's'}.`,
    duplicateStatus: (total) => `Esses arquivos já estão na lista. ${total} arquivo${total === 1 ? '' : 's'} pendente${total === 1 ? '' : 's'}.`,
    readingStatus: 'Lendo workbooks localmente no navegador...',
    mergedStatus: (rows, files, duration, saved) => `${rows.toLocaleString()} linha${rows === 1 ? '' : 's'} mesclada${rows === 1 ? '' : 's'} de ${files} arquivo${files === 1 ? '' : 's'}. Concluído em ${duration}. Economia estimada: ${saved}.`,
    noRowsStatus: 'Nenhuma linha de dados foi encontrada nos workbooks selecionados.',
    mergeError: (message) => `Não foi possível mesclar os arquivos: ${message}`,
    statFiles: 'Arquivos',
    statRows: 'Linhas mescladas',
    statColumns: 'Colunas',
    statNumeric: 'Colunas numéricas',
    statProcessTime: 'Tempo de processo',
    statTimeSaved: 'Tempo economizado',
    statsTitle: '2. Estatísticas de resumo',
    statsCopy: 'Colunas numéricas incluem contagem, soma, média, mínimo e máximo.',
    exportButton: 'Exportar summary.xlsx',
    columnHeader: 'Coluna',
    typeHeader: 'Tipo',
    nonEmptyHeader: 'Não vazio',
    blankHeader: 'Em branco',
    uniqueHeader: 'Único',
    sumHeader: 'Soma',
    averageHeader: 'Média',
    minHeader: 'Mín',
    maxHeader: 'Máx',
    statsEmpty: 'Mescle arquivos para calcular estatísticas.',
    previewTitle: '3. Prévia dos dados mesclados',
    previewCopy: 'Exibindo as primeiras 100 linhas mescladas. O workbook exportado contém todas as linhas.',
    previewEmpty: 'Ainda não há linhas mescladas.',
    faqEyebrow: 'FAQ',
    faqTitle: 'Perguntas frequentes',
    faqCopy: 'Respostas sobre privacidade, formatos, comportamento de mesclagem e alertas.',
    faqUploadQuestion: 'Meus arquivos Excel são enviados para um servidor?',
    faqUploadAnswer: 'Não. Os arquivos são lidos e mesclados localmente no navegador.',
    faqFormatsQuestion: 'Quais formatos são suportados?',
    faqFormatsAnswer: 'Você pode mesclar .xlsx, .xls e .csv.',
    faqSheetsQuestion: 'Posso mesclar várias worksheets?',
    faqSheetsAnswer: 'Sim. Você pode combinar todas as linhas em uma worksheet ou manter cada source worksheet como uma sheet separada.',
    faqMismatchQuestion: 'O que acontece se os headers forem diferentes?',
    faqMismatchAnswer: 'A ferramenta mostra um aviso com o arquivo e a sheet antes de mesclar. Você pode voltar para revisar ou continuar.',
    numericType: 'Numérico',
    textType: 'Texto/Misto',
    coverTitle: 'Mesclador local de relatórios Excel',
    coverMergeType: 'Tipo de mesclagem',
    coverOneSheet: 'Uma sheet',
    coverWorkbookSheets: 'Workbook sheets',
    coverFilesMerged: 'Arquivos mesclados',
    coverSourceSheets: 'Source worksheets',
    coverTotalRows: 'Total de linhas mescladas',
    coverProcessingTime: 'Tempo de processamento',
    coverEstimatedTimeSaved: 'Tempo economizado estimado',
    coverGeneratedAt: 'Gerado em',
    coverOutputSheet: 'Output sheet',
    coverSourceFile: 'Source file',
    coverSourceSheet: 'Source sheet',
    coverRows: 'Linhas',
    validationEyebrow: 'Revisão necessária',
    validationTitle: 'Algumas tabelas não coincidem',
    validationCopy: 'Os arquivos ainda podem ser mesclados, mas essas diferenças podem afetar o resultado. Revise antes de continuar.',
    validationHeaderIssue: 'Header diferente',
    validationTypeIssue: 'Tipo de conteúdo diferente',
    validationEmptyIssue: 'Sheet vazia ou ilegível',
    validationMore: (count) => `Mais ${count} problema${count === 1 ? '' : 's'} não exibido${count === 1 ? '' : 's'}.`,
    reviewFilesButton: 'Voltar para revisar',
    continueMergeButton: 'Continuar mesclagem',
    validationReviewStatus: 'Mesclagem pausada. Revise os arquivos pendentes e execute novamente.',
    validationContinueStatus: 'Continuando a mesclagem com as diferenças listadas...',
    missingColumns: (columns) => `Colunas ausentes: ${columns}`,
    extraColumns: (columns) => `Colunas extras: ${columns}`,
    orderMismatch: 'A ordem das colunas difere da primeira sheet.',
    typeMismatch: (column, expected, actual) => `"${column}" parece ${actual}, mas a primeira sheet parece ${expected}.`,
    emptySheetDetail: 'Nenhuma linha de dados foi encontrada nesta sheet.',
    typeNumeric: 'dados numéricos',
    typeDate: 'dados de data',
    typeText: 'dados de texto',
    typeBlank: 'dados em branco',
  },
  id: {
    ...translations.en,
    pageTitle: 'Alat Gabung Excel - Gabungkan Excel dan CSV secara lokal',
    brand: 'Penggabung Laporan Excel',
    brandTagline: 'Utilitas workbook lokal',
    languageLabel: 'Bahasa',
    shareButton: 'Salin tautan',
    shareCopied: 'Disalin',
    privacyBadge: 'File tetap lokal',
    eyebrow: 'Alat Excel di browser',
    heroTitle: 'Gabungkan laporan Excel menjadi satu workbook ringkasan',
    heroCopy: 'Tambahkan beberapa file <strong>.xlsx</strong>, <strong>.xls</strong>, atau <strong>.csv</strong>, pratinjau baris gabungan, lihat statistik, lalu ekspor <strong>summary.xlsx</strong>. Semua proses berjalan lokal di browser.',
    privacyEyebrow: 'Privasi diutamakan',
    privacyTitle: 'File Excel Anda tidak keluar dari browser ini',
    privacyCopy: 'File dibaca, diperiksa, digabungkan, dan diekspor secara lokal di perangkat Anda. Konten workbook tidak diunggah, disimpan, atau dianalisis di server.',
    privacyPointA: 'Tanpa server upload',
    privacyPointACopy: 'Data workbook diproses di browser.',
    privacyPointB: 'Tidak menyimpan file',
    privacyPointBCopy: 'Menutup halaman akan menghapus file yang dipilih.',
    privacyPointC: 'Ekspor lokal',
    privacyPointCCopy: 'Workbook akhir dibuat di perangkat Anda.',
    stepsEyebrow: 'Cara kerja',
    stepsTitle: 'Gabungkan file Excel dalam empat langkah',
    stepsCopy: 'Tambahkan file, periksa perbedaan, pilih format output, lalu ekspor workbook bersih.',
    stepOneTitle: 'Tambahkan file',
    stepOneCopy: 'Tambahkan .xlsx, .xls, atau .csv dalam satu atau beberapa batch.',
    stepTwoTitle: 'Tinjau pemeriksaan',
    stepTwoCopy: 'Lihat peringatan untuk header berbeda, sheet kosong, atau perbedaan tipe konten.',
    stepThreeTitle: 'Pilih tipe gabung',
    stepThreeCopy: 'Gabungkan semua baris ke satu sheet atau simpan source sheets dalam satu workbook.',
    stepFourTitle: 'Ekspor workbook',
    stepFourCopy: 'Unduh summary.xlsx dengan cover page, stats, dan merged data.',
    uploadTitle: '1. Upload laporan',
    uploadCopy: 'Pilih satu atau beberapa file Excel. Semua worksheets di setiap workbook akan dibaca.',
    dropTitle: 'Pilih file Excel',
    dropHint: 'atau seret dan letakkan di sini. File ditambahkan lokal dan tidak diunggah.',
    pendingLabel: 'File tertunda',
    pendingHint: 'Tambahkan file bertahap. Proses gabung mulai hanya setelah tombol Merge files diklik.',
    mergeType: 'Tipe gabung',
    oneSheet: 'Satu sheet',
    oneSheetCopy: 'Gabungkan semua baris ke satu worksheet.',
    workbookSheets: 'Workbook sheets',
    workbookSheetsCopy: 'Simpan setiap source worksheet sebagai sheet terpisah.',
    mergeButton: 'Gabungkan file',
    resetButton: 'Reset',
    waitingStatus: 'Menunggu file.',
    readyStatus: (count) => `${count} file siap digabungkan.`,
    addedStatus: (added, total) => `${added} file ditambahkan. Total ${total} file menunggu digabungkan.`,
    duplicateStatus: (total) => `File tersebut sudah ada di daftar. Total ${total} file tertunda.`,
    readingStatus: 'Membaca workbook secara lokal di browser...',
    mergedStatus: (rows, files, duration, saved) => `${rows.toLocaleString()} baris dari ${files} file berhasil digabungkan. Selesai dalam ${duration}. Estimasi waktu hemat ${saved}.`,
    noRowsStatus: 'Tidak ada baris data di workbook yang dipilih.',
    mergeError: (message) => `Tidak dapat menggabungkan file: ${message}`,
    statFiles: 'File',
    statRows: 'Baris gabungan',
    statColumns: 'Kolom',
    statNumeric: 'Kolom numerik',
    statProcessTime: 'Waktu proses',
    statTimeSaved: 'Estimasi hemat waktu',
    statsTitle: '2. Statistik ringkasan',
    statsCopy: 'Kolom numerik mencakup jumlah, total, rata-rata, minimum, dan maksimum.',
    exportButton: 'Ekspor summary.xlsx',
    columnHeader: 'Kolom',
    typeHeader: 'Tipe',
    nonEmptyHeader: 'Terisi',
    blankHeader: 'Kosong',
    uniqueHeader: 'Unik',
    sumHeader: 'Total',
    averageHeader: 'Rata-rata',
    minHeader: 'Min',
    maxHeader: 'Maks',
    statsEmpty: 'Gabungkan file untuk menghitung statistik.',
    previewTitle: '3. Pratinjau data gabungan',
    previewCopy: 'Menampilkan 100 baris pertama. Workbook ekspor berisi semua baris.',
    previewEmpty: 'Belum ada baris gabungan.',
    faqEyebrow: 'FAQ',
    faqTitle: 'Pertanyaan umum',
    faqCopy: 'Jawaban tentang privasi, format file, perilaku gabung, dan pemeriksaan mismatch.',
    faqUploadQuestion: 'Apakah file Excel saya diunggah ke server?',
    faqUploadAnswer: 'Tidak. File dibaca dan digabungkan secara lokal di browser Anda.',
    faqFormatsQuestion: 'Format file apa yang didukung?',
    faqFormatsAnswer: 'Anda dapat menggabungkan .xlsx, .xls, dan .csv.',
    faqSheetsQuestion: 'Bisakah saya menggabungkan beberapa worksheets?',
    faqSheetsAnswer: 'Bisa. Gabungkan semua baris ke satu worksheet atau simpan setiap source worksheet sebagai sheet terpisah.',
    faqMismatchQuestion: 'Apa yang terjadi jika header berbeda?',
    faqMismatchAnswer: 'Alat akan menampilkan peringatan dengan lokasi file dan sheet sebelum gabung. Anda dapat kembali memeriksa atau melanjutkan.',
    numericType: 'Numerik',
    textType: 'Teks/Campuran',
    coverTitle: 'Penggabung Laporan Excel Lokal',
    coverMergeType: 'Tipe gabung',
    coverOneSheet: 'Satu sheet',
    coverWorkbookSheets: 'Workbook sheets',
    coverFilesMerged: 'File digabungkan',
    coverSourceSheets: 'Source worksheets',
    coverTotalRows: 'Total baris gabungan',
    coverProcessingTime: 'Waktu proses',
    coverEstimatedTimeSaved: 'Estimasi waktu hemat',
    coverGeneratedAt: 'Dibuat pada',
    coverOutputSheet: 'Output sheet',
    coverSourceFile: 'Source file',
    coverSourceSheet: 'Source sheet',
    coverRows: 'Baris',
    validationEyebrow: 'Perlu ditinjau',
    validationTitle: 'Beberapa tabel tidak cocok',
    validationCopy: 'File masih bisa digabungkan, tetapi perbedaan ini dapat memengaruhi hasil. Harap tinjau terlebih dahulu.',
    validationHeaderIssue: 'Header tidak cocok',
    validationTypeIssue: 'Tipe konten tidak cocok',
    validationEmptyIssue: 'Sheet kosong atau tidak terbaca',
    validationMore: (count) => `Dan ${count} masalah lain tidak ditampilkan.`,
    reviewFilesButton: 'Kembali periksa',
    continueMergeButton: 'Lanjut gabung',
    validationReviewStatus: 'Gabung dijeda. Periksa file tertunda, lalu jalankan lagi.',
    validationContinueStatus: 'Melanjutkan gabung dengan perbedaan yang tercantum...',
    missingColumns: (columns) => `Kolom hilang: ${columns}`,
    extraColumns: (columns) => `Kolom tambahan: ${columns}`,
    orderMismatch: 'Urutan kolom berbeda dari sheet pertama.',
    typeMismatch: (column, expected, actual) => `"${column}" terlihat seperti ${actual}, tetapi sheet pertama terlihat seperti ${expected}.`,
    emptySheetDetail: 'Tidak ada baris data di sheet ini.',
    typeNumeric: 'data numerik',
    typeDate: 'data tanggal',
    typeText: 'data teks',
    typeBlank: 'data kosong',
  },
});

fileInput.addEventListener('change', (event) => {
  addSelectedFiles([...event.target.files]);
  fileInput.value = '';
});

languageSelect.addEventListener('change', (event) => {
  applyLanguage(event.target.value);
});

['dragenter', 'dragover'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add('drag-over');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove('drag-over');
  });
});

dropZone.addEventListener('drop', (event) => {
  const files = [...event.dataTransfer.files].filter(isSpreadsheetFile);
  addSelectedFiles(files);
});

mergeButton.addEventListener('click', mergeFiles);
resetButton.addEventListener('click', resetApp);
exportButton.addEventListener('click', exportWorkbook);
reviewFilesButton.addEventListener('click', closeValidationModal);
continueMergeButton.addEventListener('click', continueMergeAfterValidation);
shareButton.addEventListener('click', copyShareLink);
demoButton.addEventListener('click', loadDemoData);

applyLanguage(languageSelect.value);
incrementAnalytics('visits', 1);

function addSelectedFiles(files) {
  const existingKeys = new Set(selectedFiles.map(getFileKey));
  const filesToAdd = files.filter((file) => isSpreadsheetFile(file) && !existingKeys.has(getFileKey(file)));

  selectedFiles = [...selectedFiles, ...filesToAdd];
  renderPendingFiles();

  statusText.textContent = filesToAdd.length > 0
    ? t('addedStatus', filesToAdd.length, selectedFiles.length)
    : t('duplicateStatus', selectedFiles.length);
}

function renderPendingFiles() {
  fileList.innerHTML = '';

  selectedFiles.forEach((file) => {
    const item = document.createElement('li');
    item.innerHTML = `<span>${escapeHtml(file.name)}</span><span class="file-size">${formatBytes(file.size)}</span>`;
    fileList.appendChild(item);
  });

  pendingCount.textContent = selectedFiles.length.toLocaleString();
  mergeButton.disabled = selectedFiles.length === 0;
  resetButton.disabled = selectedFiles.length === 0 && mergedRows.length === 0;
}

async function mergeFiles() {
  if (!window.XLSX) {
    statusText.textContent = t('sheetJsError');
    return;
  }

  mergeButton.disabled = true;
  exportButton.disabled = true;
  mergeMetrics = { durationMs: 0, timeSavedMinutes: 0 };
  updateStatCards();
  statusText.textContent = t('readingStatus');

  try {
    const readStart = performance.now();
    const parsedSheets = (await Promise.all(selectedFiles.map(readWorkbook))).flat();
    const readDurationMs = performance.now() - readStart;
    const issues = findValidationIssues(parsedSheets);

    if (issues.length > 0) {
      pendingValidationSheets = parsedSheets;
      pendingValidationDurationMs = readDurationMs;
      showValidationModal(issues);
      return;
    }

    completeMerge(parsedSheets, readDurationMs);
  } catch (error) {
    console.error(error);
    statusText.textContent = t('mergeError', error.message);
  } finally {
    mergeButton.disabled = selectedFiles.length === 0;
    resetButton.disabled = selectedFiles.length === 0 && mergedRows.length === 0;
  }
}

function completeMerge(parsedSheets, readDurationMs = 0) {
  const mergeStart = performance.now();
  sourceSheets = parsedSheets;
  const headers = new Set([SOURCE_FILE, SOURCE_SHEET]);
  const rows = [];

  sourceSheets.forEach(({ fileName, sheetName, data }) => {
    data.forEach((row) => {
      Object.keys(row).forEach((header) => headers.add(header));
      rows.push({ [SOURCE_FILE]: fileName, [SOURCE_SHEET]: sheetName, ...row });
    });
  });

  allHeaders = [...headers];
  mergedRows = rows.map((row) => normalizeRow(row, allHeaders));
  summaryRows = buildSummaryRows(mergedRows, allHeaders);
  mergeMetrics = {
    durationMs: readDurationMs + performance.now() - mergeStart,
    timeSavedMinutes: estimateTimeSavedMinutes({
      fileCount: selectedFiles.length,
      sourceSheetCount: sourceSheets.length,
      rowCount: mergedRows.length,
    }),
  };

  renderStats(summaryRows);
  renderPreview(mergedRows, allHeaders);
  updateStatCards();

  exportButton.disabled = mergedRows.length === 0;
  statusText.textContent = mergedRows.length
    ? t('mergedStatus', mergedRows.length, selectedFiles.length, formatDuration(mergeMetrics.durationMs), formatTimeSaved(mergeMetrics.timeSavedMinutes))
    : t('noRowsStatus');

  if (mergedRows.length > 0) {
    incrementAnalytics('merges', 1);
    incrementAnalytics('filesProcessed', selectedFiles.length);
  }
}

function findValidationIssues(sheets) {
  const issues = [];
  const baseline = sheets.find((sheet) => sheet.headers.length > 0);

  sheets.forEach((sheet) => {
    if (sheet.data.length === 0 || sheet.headers.length === 0) {
      issues.push({
        type: 'empty',
        title: t('validationEmptyIssue'),
        location: t('issueLocation', sheet.fileName, sheet.sheetName),
        detail: t('emptySheetDetail'),
      });
    }
  });

  if (!baseline) return issues;

  const baselineHeaders = baseline.headers;
  const baselineHeaderSet = new Set(baselineHeaders);
  const baselineTypes = inferSheetColumnTypes(baseline);

  sheets.forEach((sheet) => {
    if (sheet === baseline || sheet.headers.length === 0) return;

    const headerSet = new Set(sheet.headers);
    const missingColumns = baselineHeaders.filter((header) => !headerSet.has(header));
    const extraColumns = sheet.headers.filter((header) => !baselineHeaderSet.has(header));
    const sameColumnSet = missingColumns.length === 0 && extraColumns.length === 0;
    const orderDiffers = sameColumnSet && sheet.headers.join('\u0001') !== baselineHeaders.join('\u0001');
    const headerDetails = [];

    if (missingColumns.length > 0) headerDetails.push(t('missingColumns', missingColumns.join(', ')));
    if (extraColumns.length > 0) headerDetails.push(t('extraColumns', extraColumns.join(', ')));
    if (orderDiffers) headerDetails.push(t('orderMismatch'));

    if (headerDetails.length > 0) {
      issues.push({
        type: 'header',
        title: t('validationHeaderIssue'),
        location: t('issueLocation', sheet.fileName, sheet.sheetName),
        detail: headerDetails.join(' '),
      });
    }

    const sheetTypes = inferSheetColumnTypes(sheet);
    baselineHeaders
      .filter((header) => headerSet.has(header))
      .forEach((header) => {
        const expectedType = baselineTypes[header];
        const actualType = sheetTypes[header];

        if (expectedType && actualType && expectedType !== 'blank' && actualType !== 'blank' && expectedType !== actualType) {
          issues.push({
            type: 'type',
            title: t('validationTypeIssue'),
            location: t('issueLocation', sheet.fileName, sheet.sheetName),
            detail: t('typeMismatch', header, localizeInferredType(expectedType), localizeInferredType(actualType)),
          });
        }
      });
  });

  return issues;
}

function showValidationModal(issues) {
  validationList.innerHTML = '';

  issues.slice(0, 20).forEach((issue) => {
    const item = document.createElement('article');
    item.className = 'validation-item';
    item.innerHTML = `
      <strong>${escapeHtml(issue.title)}</strong>
      <span>${escapeHtml(issue.location)}</span>
      <small>${escapeHtml(issue.detail)}</small>
    `;
    validationList.appendChild(item);
  });

  if (issues.length > 20) {
    const item = document.createElement('article');
    item.className = 'validation-item';
    item.innerHTML = `<strong>${escapeHtml(t('validationMore', issues.length - 20))}</strong>`;
    validationList.appendChild(item);
  }

  validationModal.hidden = false;
  continueMergeButton.focus();
}

function closeValidationModal() {
  validationModal.hidden = true;
  pendingValidationSheets = [];
  pendingValidationDurationMs = 0;
  statusText.textContent = t('validationReviewStatus');
}

function continueMergeAfterValidation() {
  validationModal.hidden = true;
  statusText.textContent = t('validationContinueStatus');
  const sheets = pendingValidationSheets;
  const readDurationMs = pendingValidationDurationMs;
  pendingValidationSheets = [];
  pendingValidationDurationMs = 0;
  completeMerge(sheets, readDurationMs);
  mergeButton.disabled = selectedFiles.length === 0;
  resetButton.disabled = selectedFiles.length === 0 && mergedRows.length === 0;
}

function inferSheetColumnTypes(sheet) {
  return sheet.headers.reduce((types, header) => {
    const values = sheet.data.map((row) => row[header]).filter((value) => String(value ?? '').trim() !== '');
    types[header] = inferValuesType(values);
    return types;
  }, {});
}

function inferValuesType(values) {
  if (values.length === 0) return 'blank';

  let numericCount = 0;
  let dateCount = 0;

  values.forEach((value) => {
    if (value instanceof Date && Number.isFinite(value.getTime())) {
      dateCount += 1;
      return;
    }

    const text = String(value).trim();
    const numericValue = Number(text.replace(/[$,%\s,]/g, ''));
    if (text !== '' && Number.isFinite(numericValue)) {
      numericCount += 1;
      return;
    }

    if (!Number.isNaN(Date.parse(text)) && /\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(text)) {
      dateCount += 1;
    }
  });

  if (numericCount === values.length) return 'numeric';
  if (dateCount === values.length) return 'date';
  return 'text';
}

function localizeInferredType(type) {
  return {
    numeric: t('typeNumeric'),
    date: t('typeDate'),
    text: t('typeText'),
    blank: t('typeBlank'),
  }[type] || type;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
        if (workbook.SheetNames.length === 0) {
          resolve([{ fileName: file.name, sheetName: 'No sheets', headers: [], data: [] }]);
          return;
        }

        const sheets = workbook.SheetNames.map((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
          const dedupedData = dedupeObjectKeys(data);
          const headers = [...new Set(dedupedData.flatMap((row) => Object.keys(row)))];
          return { fileName: file.name, sheetName, headers, data: dedupedData };
        });

        resolve(sheets);
      } catch (error) {
        reject(new Error(`${file.name}: ${error.message}`));
      }
    };

    reader.onerror = () => reject(new Error(`${file.name}: ${reader.error?.message || 'file read failed'}`));
    reader.readAsArrayBuffer(file);
  });
}

function buildSummaryRows(rows, headers) {
  return headers.map((header) => {
    const values = rows.map((row) => row[header]);
    const nonEmptyValues = values.filter((value) => String(value).trim() !== '');
    const numericValues = nonEmptyValues
      .map((value) => Number(String(value).replace(/[$,%\s,]/g, '')))
      .filter((value) => Number.isFinite(value));
    const isNumeric = numericValues.length > 0 && numericValues.length === nonEmptyValues.length;
    const sum = numericValues.reduce((total, value) => total + value, 0);

    return {
      Column: header,
      Type: isNumeric ? 'Numeric' : 'Text/Mixed',
      'Non-empty': nonEmptyValues.length,
      Blank: values.length - nonEmptyValues.length,
      Unique: new Set(nonEmptyValues.map((value) => String(value).trim())).size,
      Sum: isNumeric ? round(sum) : '',
      Average: isNumeric ? round(sum / numericValues.length) : '',
      Min: isNumeric ? round(Math.min(...numericValues)) : '',
      Max: isNumeric ? round(Math.max(...numericValues)) : '',
    };
  });
}

function renderStats(rows) {
  statsTableBody.innerHTML = '';

  if (rows.length === 0) {
    statsTableBody.innerHTML = `<tr><td colspan="9" class="empty-state">${t('statsEmpty')}</td></tr>`;
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    Object.entries(row).forEach(([key, value]) => {
      const td = document.createElement('td');
      td.textContent = key === 'Type' ? localizeColumnType(value) : value;
      tr.appendChild(td);
    });
    statsTableBody.appendChild(tr);
  });
}

function renderPreview(rows, headers) {
  previewTable.innerHTML = '';

  if (rows.length === 0) {
    previewTable.innerHTML = `<tbody><tr><td class="empty-state">${t('previewEmpty')}</td></tr></tbody>`;
    return;
  }

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  headers.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement('tbody');
  rows.slice(0, 100).forEach((row) => {
    const tr = document.createElement('tr');
    headers.forEach((header) => {
      const td = document.createElement('td');
      td.textContent = row[header];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewTable.append(thead, tbody);
}

function exportWorkbook() {
  if (!window.XLSX || mergedRows.length === 0) return;

  const workbook = XLSX.utils.book_new();
  const mergeMode = getMergeMode();
  const outputSheets = mergeMode === 'workbook'
    ? buildWorkbookSheets(sourceSheets)
    : [{ name: 'Merged Data', rows: mergedRows, headers: allHeaders }];
  const coverSheet = buildCoverSheet({
    mergeMode,
    fileCount: selectedFiles.length,
    totalRows: mergedRows.length,
    sourceSheetCount: sourceSheets.length,
    outputSheets,
    durationText: formatDuration(mergeMetrics.durationMs),
    timeSavedText: formatTimeSaved(mergeMetrics.timeSavedMinutes),
  });

  XLSX.utils.book_append_sheet(workbook, coverSheet, 'Cover Page');

  if (mergeMode === 'workbook') {
    outputSheets.forEach((sheet) => {
      const worksheet = XLSX.utils.json_to_sheet(sheet.rows, { header: sheet.headers });
      XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
    });
  } else {
    const mergedSheet = XLSX.utils.json_to_sheet(mergedRows, { header: allHeaders });
    XLSX.utils.book_append_sheet(workbook, mergedSheet, 'Merged Data');
  }

  const statsSheet = XLSX.utils.json_to_sheet(summaryRows);

  XLSX.utils.book_append_sheet(workbook, statsSheet, 'Summary Stats');
  XLSX.writeFile(workbook, 'summary.xlsx');
  incrementAnalytics('exportsGenerated', 1);
}

function resetApp() {
  selectedFiles = [];
  sourceSheets = [];
  mergedRows = [];
  summaryRows = [];
  allHeaders = [];
  pendingValidationSheets = [];
  pendingValidationDurationMs = 0;
  mergeMetrics = { durationMs: 0, timeSavedMinutes: 0 };
  validationModal.hidden = true;
  fileInput.value = '';
  fileList.innerHTML = '';
  pendingCount.textContent = '0';
  mergeButton.disabled = true;
  resetButton.disabled = true;
  exportButton.disabled = true;
  statusText.textContent = t('waitingStatus');
  renderStats([]);
  renderPreview([], []);
  updateStatCards();
}

function updateStatCards() {
  statFiles.textContent = selectedFiles.length.toLocaleString();
  statRows.textContent = mergedRows.length.toLocaleString();
  statColumns.textContent = allHeaders.length.toLocaleString();
  statNumeric.textContent = summaryRows.filter((row) => row.Type === 'Numeric').length.toLocaleString();
  statProcessTime.textContent = formatDuration(mergeMetrics.durationMs);
  statTimeSaved.textContent = formatTimeSaved(mergeMetrics.timeSavedMinutes);
  statSheets.textContent = sourceSheets.length.toLocaleString();
}

function loadDemoData() {
  if (!window.XLSX) {
    statusText.textContent = t('sheetJsError');
    return;
  }

  const demoSheets = buildDemoSheets();
  selectedFiles = demoSheets.map((sheet) => createDemoFile(sheet));
  pendingValidationSheets = [];
  pendingValidationDurationMs = 0;
  mergeMetrics = { durationMs: 0, timeSavedMinutes: 0 };

  renderPendingFiles();
  statusText.textContent = t('demoLoadedStatus');
  completeMerge(demoSheets, 0);
  document.querySelector('#stats-title')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function buildDemoSheets() {
  const months = [
    { fileName: 'Sales_Report_Jan.xlsx', sheetName: 'Sheet_Jan', month: 'Jan', base: 1200 },
    { fileName: 'Sales_Report_Feb.xlsx', sheetName: 'Sheet_Feb', month: 'Feb', base: 1320 },
    { fileName: 'Sales_Report_Mar.xlsx', sheetName: 'Sheet_Mar', month: 'Mar', base: 1480 },
  ];
  const stores = ['North Store', 'South Store', 'Online Store', 'Outlet Store', 'Airport Store'];

  return months.map((month) => {
    const data = stores.map((store, index) => ({
      Month: month.month,
      Store: store,
      Orders: 42 + index * 7 + month.month.length,
      Revenue: month.base + index * 215,
      Manager: ['Ava', 'Ben', 'Chloe', 'Daniel', 'Emma'][index],
    }));

    return {
      fileName: month.fileName,
      sheetName: month.sheetName,
      headers: ['Month', 'Store', 'Orders', 'Revenue', 'Manager'],
      data,
    };
  });
}

function createDemoFile(sheet) {
  if (window.XLSX && typeof File !== 'undefined') {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(sheet.data, { header: sheet.headers });
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.sheetName);
    const content = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    return new File([content], sheet.fileName, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      lastModified: Date.now(),
    });
  }

  return {
    name: sheet.fileName,
    size: 18432,
    lastModified: Date.now(),
  };
}

function estimateTimeSavedMinutes({ fileCount, sourceSheetCount, rowCount }) {
  if (rowCount === 0) return 0;

  const fileSetupMinutes = fileCount * 2;
  const sheetReviewMinutes = sourceSheetCount;
  const rowHandlingMinutes = Math.ceil(rowCount / 100);

  return Math.max(1, fileSetupMinutes + sheetReviewMinutes + rowHandlingMinutes);
}

function formatDuration(milliseconds) {
  if (!milliseconds || milliseconds <= 0) {
    return formatDurationUnit(0, 'seconds');
  }

  const seconds = milliseconds / 1000;

  if (seconds < 60) {
    const value = seconds < 10 ? seconds.toFixed(1) : Math.round(seconds).toString();
    return formatDurationUnit(value, 'seconds');
  }

  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = Math.round(seconds % 60);
  return `${formatDurationUnit(minutes, 'minutes')} ${formatDurationUnit(remainingSeconds, 'seconds')}`;
}

function formatTimeSaved(minutes) {
  if (!minutes || minutes <= 0) {
    return formatDurationUnit(0, 'minutes');
  }

  if (minutes < 60) {
    return formatDurationUnit(minutes, 'minutes');
  }

  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;

  return remainingMinutes > 0
    ? `${formatDurationUnit(hours, 'hours')} ${formatDurationUnit(remainingMinutes, 'minutes')}`
    : formatDurationUnit(hours, 'hours');
}

function formatDurationUnit(value, unit) {
  const unitLabels = {
    en: { seconds: 's', minutes: 'min', hours: value === 1 ? 'hr' : 'hrs' },
    zh: { seconds: '秒', minutes: '分钟', hours: '小时' },
    hi: { seconds: 'सेकंड', minutes: 'मिनट', hours: 'घंटे' },
    ja: { seconds: '秒', minutes: '分', hours: '時間' },
    pt: { seconds: 's', minutes: 'min', hours: 'h' },
    id: { seconds: 'dtk', minutes: 'mnt', hours: 'jam' },
  };
  const labels = unitLabels[currentLanguage] || unitLabels.en;

  if (currentLanguage === 'zh' || currentLanguage === 'ja') {
    return `${value}${labels[unit]}`;
  }

  return `${value} ${labels[unit]}`;
}

function getAnalyticsStats() {
  try {
    const stored = JSON.parse(localStorage.getItem(ANALYTICS_KEY) || '{}');
    return {
      visits: Number(stored.visits) || 0,
      merges: Number(stored.merges) || 0,
      filesProcessed: Number(stored.filesProcessed) || 0,
      exportsGenerated: Number(stored.exportsGenerated) || 0,
    };
  } catch (error) {
    return { visits: 0, merges: 0, filesProcessed: 0, exportsGenerated: 0 };
  }
}

function saveAnalyticsStats(stats) {
  localStorage.setItem(ANALYTICS_KEY, JSON.stringify(stats));
}

function incrementAnalytics(key, amount) {
  const stats = getAnalyticsStats();
  stats[key] = (Number(stats[key]) || 0) + amount;
  saveAnalyticsStats(stats);
  renderAnalyticsStats(stats);
}

function renderAnalyticsStats(stats = getAnalyticsStats()) {
  analyticsVisits.textContent = stats.visits.toLocaleString();
  analyticsMerges.textContent = stats.merges.toLocaleString();
  analyticsFiles.textContent = stats.filesProcessed.toLocaleString();
  analyticsExports.textContent = stats.exportsGenerated.toLocaleString();
}

function buildCoverSheet({ mergeMode, fileCount, totalRows, sourceSheetCount, outputSheets, durationText, timeSavedText }) {
  const rows = [
    [t('coverTitle')],
    [],
    [t('coverMergeType'), mergeMode === 'workbook' ? t('coverWorkbookSheets') : t('coverOneSheet')],
    [t('coverFilesMerged'), fileCount],
    [t('coverSourceSheets'), sourceSheetCount],
    [t('coverTotalRows'), totalRows],
    [t('coverProcessingTime'), durationText],
    [t('coverEstimatedTimeSaved'), timeSavedText],
    [t('coverGeneratedAt'), new Date().toLocaleString()],
    [],
  ];

  if (mergeMode === 'workbook') {
    rows.push([t('coverOutputSheet'), t('coverSourceFile'), t('coverSourceSheet'), t('coverRows')]);
    outputSheets.forEach((sheet) => {
      rows.push([sheet.name, sheet.fileName, sheet.sourceSheetName, sheet.rows.length]);
    });
  } else {
    rows.push([t('coverOutputSheet'), t('coverRows')]);
    outputSheets.forEach((sheet) => {
      rows.push([sheet.name, sheet.rows.length]);
    });
    rows.push(['Summary Stats', summaryRows.length]);
  }

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet['!cols'] = [{ wch: 28 }, { wch: 28 }, { wch: 28 }, { wch: 14 }];

  if (mergeMode === 'workbook') {
    const firstLinkedSheetRow = rows.findIndex((row) => row[0] === t('coverOutputSheet')) + 1;
    outputSheets.forEach((sheet, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: firstLinkedSheetRow + index, c: 0 });
      worksheet[cellAddress].l = {
        Target: `#'${sheet.name.replace(/'/g, "''")}'!A1`,
        Tooltip: `Open ${sheet.name}`,
      };
      worksheet[cellAddress].s = { font: { color: { rgb: '0563C1' }, underline: true } };
    });
  }

  return worksheet;
}

function buildWorkbookSheets(sheets) {
  const usedNames = new Set(['Cover Page', 'Summary Stats']);

  return sheets.map((sheet) => {
    const headers = [...new Set([SOURCE_FILE, SOURCE_SHEET, ...sheet.headers])];
    const rows = sheet.data.map((row) => normalizeRow({
      [SOURCE_FILE]: sheet.fileName,
      [SOURCE_SHEET]: sheet.sheetName,
      ...row,
    }, headers));

    return {
      name: getUniqueSheetName(sheet.sheetName, usedNames),
      fileName: sheet.fileName,
      sourceSheetName: sheet.sheetName,
      headers,
      rows,
    };
  });
}

function getMergeMode() {
  return document.querySelector('input[name="merge-mode"]:checked')?.value || 'single';
}

function applyLanguage(language) {
  currentLanguage = translations[language] ? language : 'en';
  document.documentElement.lang = languageAttributes[currentLanguage] || 'en';
  document.title = t('pageTitle');

  document.querySelectorAll('[data-i18n]').forEach((element) => {
    element.textContent = t(element.dataset.i18n);
  });

  document.querySelectorAll('[data-i18n-html]').forEach((element) => {
    element.innerHTML = t(element.dataset.i18nHtml);
  });

  if (selectedFiles.length > 0) {
    statusText.textContent = mergedRows.length
      ? t('mergedStatus', mergedRows.length, selectedFiles.length, formatDuration(mergeMetrics.durationMs), formatTimeSaved(mergeMetrics.timeSavedMinutes))
      : t('readyStatus', selectedFiles.length);
  } else {
    statusText.textContent = t('waitingStatus');
  }

  renderStats(summaryRows);
  renderPreview(mergedRows, allHeaders);
  updateStatCards();
}

async function copyShareLink() {
  const originalLabel = t('shareButton');

  try {
    await navigator.clipboard.writeText(SITE_URL);
    shareButton.textContent = t('shareCopied');
    setTimeout(() => {
      shareButton.textContent = originalLabel;
    }, 1600);
  } catch (error) {
    shareButton.textContent = SITE_URL;
  }
}

function t(key, ...args) {
  const value = translations[currentLanguage][key] ?? translations.en[key] ?? key;
  return typeof value === 'function' ? value(...args) : value;
}

function localizeColumnType(value) {
  if (value === 'Numeric') return t('numericType');
  if (value === 'Text/Mixed') return t('textType');
  return value;
}

function getUniqueSheetName(name, usedNames) {
  const baseName = sanitizeSheetName(name);
  let candidate = baseName;
  let index = 2;

  while (usedNames.has(candidate)) {
    const suffix = ` ${index}`;
    candidate = `${baseName.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }

  usedNames.add(candidate);
  return candidate;
}

function sanitizeSheetName(name) {
  const cleaned = String(name || 'Sheet')
    .replace(/[\][*?/\\:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return (cleaned || 'Sheet').slice(0, 31);
}

function normalizeRow(row, headers) {
  return headers.reduce((normalized, header) => {
    normalized[header] = row[header] ?? '';
    return normalized;
  }, {});
}

function dedupeObjectKeys(rows) {
  return rows.map((row) => {
    const deduped = {};
    Object.entries(row).forEach(([key, value]) => {
      const cleanKey = String(key).trim() || 'Unnamed Column';
      let finalKey = cleanKey;
      let index = 2;

      while (Object.prototype.hasOwnProperty.call(deduped, finalKey)) {
        finalKey = `${cleanKey} ${index}`;
        index += 1;
      }

      deduped[finalKey] = value;
    });
    return deduped;
  });
}

function createFileList(files) {
  const dataTransfer = new DataTransfer();
  files.forEach((file) => dataTransfer.items.add(file));
  return dataTransfer.files;
}

function getFileKey(file) {
  return `${file.name}-${file.size}-${file.lastModified}`;
}

function isSpreadsheetFile(file) {
  return /\.(xlsx|xls|csv)$/i.test(file.name);
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function round(value) {
  return Number.isFinite(value) ? Math.round((value + Number.EPSILON) * 100) / 100 : '';
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (character) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;',
  }[character]));
}
