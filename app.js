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
const languageSelect = document.querySelector('#language-select');
const pendingCount = document.querySelector('#pending-count');
const validationModal = document.querySelector('#validation-modal');
const validationList = document.querySelector('#validation-list');
const reviewFilesButton = document.querySelector('#review-files-button');
const continueMergeButton = document.querySelector('#continue-merge-button');
const shareButton = document.querySelector('#share-button');

let selectedFiles = [];
let sourceSheets = [];
let mergedRows = [];
let summaryRows = [];
let allHeaders = [];
let currentLanguage = 'en';
let pendingValidationSheets = [];

const SOURCE_FILE = 'Source File';
const SOURCE_SHEET = 'Source Sheet';
const SITE_URL = 'https://filetools-lab.github.io/excelmerge/';
const translations = {
  en: {
    pageTitle: 'Local Excel Report Merger',
    brand: 'Excel Report Merger',
    brandTagline: 'Local workbook utility',
    languageLabel: 'Language',
    shareButton: 'Copy link',
    shareCopied: 'Copied',
    feedbackLink: 'Feedback',
    privacyBadge: 'Files stay local',
    eyebrow: 'Browser-only Excel utility',
    heroTitle: 'Merge Excel reports into one summary workbook',
    heroCopy: 'Upload multiple <strong>.xlsx</strong>, <strong>.xls</strong>, or <strong>.csv</strong> files, preview the combined rows, review summary statistics, and export <strong>summary.xlsx</strong>. Processing happens locally in your browser.',
    privacyEyebrow: 'Privacy-first processing',
    privacyTitle: 'Your Excel files never leave this browser',
    privacyCopy: 'Files are read, checked, merged, and exported locally on your device. This tool does not upload, store, or analyze your workbook contents on a server.',
    privacyPointA: 'No upload server',
    privacyPointACopy: 'Workbook data is processed in the browser.',
    privacyPointB: 'No saved files',
    privacyPointBCopy: 'Closing the page clears the selected files.',
    privacyPointC: 'Export locally',
    privacyPointCCopy: 'The final workbook is generated on your device.',
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
    uploadCopy: 'Select one or more Excel files. Every worksheet in each workbook is merged.',
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
    mergedStatus: (rows, files) => `Merged ${rows.toLocaleString()} row${rows === 1 ? '' : 's'} from ${files} file${files === 1 ? '' : 's'}.`,
    noRowsStatus: 'No data rows were found in the selected workbooks.',
    mergeError: (message) => `Unable to merge files: ${message}`,
    statFiles: 'Files',
    statRows: 'Merged rows',
    statColumns: 'Columns',
    statNumeric: 'Numeric columns',
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
    numericType: 'Numeric',
    textType: 'Text/Mixed',
    coverTitle: 'Local Excel Report Merger',
    coverMergeType: 'Merge type',
    coverOneSheet: 'One sheet',
    coverWorkbookSheets: 'Workbook sheets',
    coverFilesMerged: 'Files merged',
    coverSourceSheets: 'Source worksheets merged',
    coverTotalRows: 'Total merged rows',
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
    pageTitle: '本地 Excel 报表合并工具',
    brand: 'Excel 报表合并工具',
    brandTagline: '本地工作簿处理工具',
    languageLabel: '语言',
    shareButton: '复制链接',
    shareCopied: '已复制',
    feedbackLink: '反馈',
    privacyBadge: '文件仅保留在本地',
    eyebrow: '仅在浏览器本地处理的 Excel 工具',
    heroTitle: '将多个 Excel 报表合并成一个汇总工作簿',
    heroCopy: '上传多个 <strong>.xlsx</strong>、<strong>.xls</strong> 或 <strong>.csv</strong> 文件，预览合并后的数据，查看汇总统计，并导出 <strong>summary.xlsx</strong>。所有处理都在你的浏览器本地完成。',
    privacyEyebrow: '隐私优先',
    privacyTitle: '你的 Excel 文件不会离开当前浏览器',
    privacyCopy: '文件读取、检查、合并和导出都在你的设备本地完成。这个工具不会把工作簿内容上传到服务器，也不会存储或分析你的表格数据。',
    privacyPointA: '没有上传服务器',
    privacyPointACopy: '工作簿数据直接在浏览器里处理。',
    privacyPointB: '不保存文件',
    privacyPointBCopy: '关闭页面后，已选择的文件会被清空。',
    privacyPointC: '本地导出',
    privacyPointCCopy: '最终工作簿在你的设备上生成。',
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
    uploadCopy: '选择一个或多个 Excel 文件。每个工作簿里的所有 worksheet 都会被读取。',
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
    mergedStatus: (rows, files) => `已从 ${files} 个文件中合并 ${rows.toLocaleString()} 行数据。`,
    noRowsStatus: '所选工作簿中没有找到数据行。',
    mergeError: (message) => `无法合并文件：${message}`,
    statFiles: '文件数',
    statRows: '合并行数',
    statColumns: '列数',
    statNumeric: '数字列',
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
    numericType: '数字',
    textType: '文本/混合',
    coverTitle: '本地 Excel 报表合并工具',
    coverMergeType: '合并类型',
    coverOneSheet: '合并到一个 sheet',
    coverWorkbookSheets: '合并到多个 sheet',
    coverFilesMerged: '合并文件数',
    coverSourceSheets: 'Sheet 数',
    coverTotalRows: '合并总行数',
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

applyLanguage(languageSelect.value);

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
  statusText.textContent = t('readingStatus');

  try {
    const parsedSheets = (await Promise.all(selectedFiles.map(readWorkbook))).flat();
    const issues = findValidationIssues(parsedSheets);

    if (issues.length > 0) {
      pendingValidationSheets = parsedSheets;
      showValidationModal(issues);
      return;
    }

    completeMerge(parsedSheets);
  } catch (error) {
    console.error(error);
    statusText.textContent = t('mergeError', error.message);
  } finally {
    mergeButton.disabled = selectedFiles.length === 0;
    resetButton.disabled = selectedFiles.length === 0 && mergedRows.length === 0;
  }
}

function completeMerge(parsedSheets) {
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

  renderStats(summaryRows);
  renderPreview(mergedRows, allHeaders);
  updateStatCards();

  exportButton.disabled = mergedRows.length === 0;
  statusText.textContent = mergedRows.length
    ? t('mergedStatus', mergedRows.length, selectedFiles.length)
    : t('noRowsStatus');
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
  statusText.textContent = t('validationReviewStatus');
}

function continueMergeAfterValidation() {
  validationModal.hidden = true;
  statusText.textContent = t('validationContinueStatus');
  completeMerge(pendingValidationSheets);
  pendingValidationSheets = [];
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
}

function resetApp() {
  selectedFiles = [];
  sourceSheets = [];
  mergedRows = [];
  summaryRows = [];
  allHeaders = [];
  pendingValidationSheets = [];
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
}

function buildCoverSheet({ mergeMode, fileCount, totalRows, sourceSheetCount, outputSheets }) {
  const rows = [
    [t('coverTitle')],
    [],
    [t('coverMergeType'), mergeMode === 'workbook' ? t('coverWorkbookSheets') : t('coverOneSheet')],
    [t('coverFilesMerged'), fileCount],
    [t('coverSourceSheets'), sourceSheetCount],
    [t('coverTotalRows'), totalRows],
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
    outputSheets.forEach((sheet, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 8 + index, c: 0 });
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
  document.documentElement.lang = currentLanguage === 'zh' ? 'zh-CN' : 'en';
  document.title = t('pageTitle');

  document.querySelectorAll('[data-i18n]').forEach((element) => {
    element.textContent = t(element.dataset.i18n);
  });

  document.querySelectorAll('[data-i18n-html]').forEach((element) => {
    element.innerHTML = t(element.dataset.i18nHtml);
  });

  if (selectedFiles.length > 0) {
    statusText.textContent = mergedRows.length
      ? t('mergedStatus', mergedRows.length, selectedFiles.length)
      : t('readyStatus', selectedFiles.length);
  } else {
    statusText.textContent = t('waitingStatus');
  }

  renderStats(summaryRows);
  renderPreview(mergedRows, allHeaders);
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
