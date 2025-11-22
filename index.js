/**
 * index.js (scheduling update)
 * 使用 exceljs 讀寫本機 Excel 並透過 YouTube Data API 上傳影片（可排程）
 *
 * 這個版本新增更友善的排程選項：
 * - 立即公開（default）
 * - 起始時間 + 每支影片間隔（start + interval）
 * - 每日固定時間發佈 N 支影片（daily schedule），可設定起始日期、每天幾支、以及同日內影片之間的間隔分鐘數
 *
 * 需求：
 *  - credentials.json (Google OAuth desktop client) 放在專案根目錄 (若要上傳)
 *  - npm install
 *  - node index.js
 *
 * 已包含先前對 inquirer.createPromptModule 的容錯處理。
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const inquirer = require('inquirer');
const open = require('open');
const ExcelJS = require('exceljs');

// robust prompt init (works with multiple inquirer builds)
let prompt;
try {
  if (typeof inquirer.createPromptModule === 'function') {
    prompt = inquirer.createPromptModule();
  } else if (typeof inquirer.prompt === 'function') {
    prompt = (...args) => inquirer.prompt(...args);
  } else {
    throw new Error('inquirer has no createPromptModule or prompt export');
  }
} catch (e) {
  throw new Error(
    '無法初始化 inquirer 的 prompt 函式。請檢查是否有本地檔案或資料夾名稱為 inquirer，或 node_modules 是否正確安裝。錯誤: ' +
    e.message
  );
}

const SCOPES = [
  'https://www.googleapis.com/auth/youtube.upload',
  'https://www.googleapis.com/auth/youtube'
];

const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');
const TOKEN_PATH = path.join(process.cwd(), 'token.json');

async function loadCredentials() {
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new Error(`找不到 ${CREDENTIALS_PATH}。請至 Google Cloud 建立 OAuth Client (Desktop)，下載 credentials.json 並放到專案根目錄。`);
  }
  return JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));
}

async function authorize() {
  const credentials = await loadCredentials();
  const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
  });
  console.log('請於瀏覽器中開啟以下網址並授權：');
  console.log(authUrl);
  try { await open(authUrl); } catch (e) { /* ignore */ }

  const answers = await prompt([
    { name: 'code', message: '授權完成後，請貼上授權碼（或若使用自動 redirect 則直接按 Enter）:', default: '' }
  ]);
  const code = (answers && answers.code) ? answers.code.trim() : '';
  if (code) {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
    console.log(`已把 token 存到 ${TOKEN_PATH}`);
    return oAuth2Client;
  } else {
    throw new Error('需要授權碼來完成 OAuth2 流程。');
  }
}

function normalizeHeader(h) {
  if (!h && h !== 0) return '';
  return h.toString().trim().toLowerCase();
}

function findFileByBase(folderPath, baseName) {
  try {
    const files = fs.readdirSync(folderPath);
    const exact = files.find(f => path.parse(f).name === baseName);
    if (exact) return path.join(folderPath, exact);
    const starts = files.find(f => path.parse(f).name.startsWith(baseName));
    if (starts) return path.join(folderPath, starts);
    return null;
  } catch (e) {
    return null;
  }
}

async function readSheetToArrayOfArrays(excelPath, sheetName) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new Error(`工作表 ${sheetName} 不存在。`);
  const values = [];
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const rowVals = [];
    for (let c = 1; c <= Math.max(row.actualCellCount, row.cellCount); c++) {
      const cell = row.getCell(c);
      let v = '';
      if (cell && (cell.value !== null && cell.value !== undefined)) {
        if (typeof cell.text === 'string' && cell.text !== '') v = cell.text;
        else if (typeof cell.value === 'object' && cell.value.richText) {
          v = cell.value.richText.map(t => t.text).join('');
        } else {
          v = cell.value.toString();
        }
      } else {
        v = '';
      }
      rowVals[c - 1] = v;
    }
    values.push(rowVals);
  });
  return { workbook, worksheet, values };
}

async function writeArrayOfArraysToSheetAndSave(workbook, worksheet, values, excelPath) {
  if (worksheet.rowCount > 0) {
    worksheet.spliceRows(1, worksheet.rowCount);
  }
  for (let r = 0; r < values.length; r++) {
    const row = Array.isArray(values[r]) ? values[r] : [values[r]];
    worksheet.addRow(row);
  }
  await workbook.xlsx.writeFile(excelPath);
}

/**
 * Compute publishAt (ISO string) for the next scheduled video, based on scheduleOptions.
 * scheduledIndex: zero-based index of this upload among uploads that receive scheduling (0,1,2,...)
 *
 * scheduleOptions structure:
 * - type: 'none' | 'start_interval' | 'daily'
 * - for start_interval:
 *    { type:'start_interval', startAt: Date, intervalMins: number }
 * - for daily:
 *    { type:'daily', startDate: 'YYYY-MM-DD', timeOfDay: 'HH:MM' (24h), perDay: number, spacingMins: number }
 */
function computePublishAtForIndex(scheduledIndex, scheduleOptions) {
  if (!scheduleOptions || scheduleOptions.type === 'none') return null;
  if (scheduleOptions.type === 'start_interval') {
    const dt = new Date(scheduleOptions.startAt.getTime() + scheduledIndex * scheduleOptions.intervalMins * 60 * 1000);
    return dt.toISOString();
  }
  if (scheduleOptions.type === 'daily') {
    const { startDate, timeOfDay, perDay, spacingMins } = scheduleOptions;
    // parse startDate + timeOfDay as local time
    const [hourStr, minStr] = timeOfDay.split(':');
    const [year, month, day] = startDate.split('-').map(Number);
    const base = new Date(year, month - 1, day, Number(hourStr), Number(minStr), 0, 0); // local
    const dayOffset = Math.floor(scheduledIndex / perDay);
    const indexInDay = scheduledIndex % perDay;
    // add dayOffset days
    const dt = new Date(base.getTime() + dayOffset * 24 * 60 * 60 * 1000 + indexInDay * spacingMins * 60 * 1000);
    return dt.toISOString();
  }
  return null;
}

async function main() {
  try {
    const basicAnswers = await prompt([
      {
        type: 'input',
        name: 'excelPath',
        message: '請輸入要讀取的 Excel 檔案路徑 (.xlsx)：',
        default: './data.xlsx',
        validate: (input) => {
          if (!input) return '請輸入檔案路徑';
          const p = path.resolve(input);
          if (!fs.existsSync(p)) return '檔案不存在，請確認路徑';
          if (!p.toLowerCase().endsWith('.xlsx')) return '僅支援 .xlsx 檔';
          return true;
        }
      },
      {
        name: 'sheetName',
        message: '請輸入要讀取的 worksheet 名稱：',
        default: 'Output_100'
      },
      {
        name: 'idColumn',
        message: '請輸入用來對應影片檔名的欄位名稱（例如 編號）：',
        default: '編號'
      },
      {
        name: 'folderPath',
        message: '請輸入影片資料夾的絕對或相對路徑：',
        default: './videos'
      },
      {
        type: 'confirm',
        name: 'doUpload',
        message: '你要程式自動上傳到 YouTube 嗎？（若選否，將只讀取 Excel 並輸出 Uploaded 欄位為模擬結果）',
        default: true
      }
    ]);

    // scheduling strategy selection
    const scheduleChoice = await prompt([
      {
        type: 'list',
        name: 'scheduleType',
        message: '選擇影片發布（publishAt）排程方式：',
        choices: [
          { name: '立即公開（不排程）', value: 'none' },
          { name: '指定起始時間，之後每支影片間隔固定分鐘數', value: 'start_interval' },
          { name: '每天固定時間發佈 N 支影片（可設定同日內間隔）', value: 'daily' }
        ],
        default: 'none'
      }
    ]);

    let scheduleOptions = { type: 'none' };
    if (scheduleChoice.scheduleType === 'start_interval') {
      const s = await prompt([
        {
          name: 'startAt',
          message: '排程起始時間 (ISO 8601，本地或帶Z皆可，例如 2025-11-22T10:00:00 或 2025-11-22T10:00:00Z)：',
          default: new Date(Date.now() + 5 * 60 * 1000).toISOString().replace(/\.\d+Z$/, 'Z')
        },
        {
          name: 'intervalMins',
          message: '每支影片間隔 (分鐘)：',
          default: 10,
          validate: v => (!isNaN(Number(v)) && Number(v) >= 0) ? true : '請輸入數字'
        }
      ]);
      scheduleOptions = {
        type: 'start_interval',
        startAt: new Date(s.startAt),
        intervalMins: Number(s.intervalMins)
      };
    } else if (scheduleChoice.scheduleType === 'daily') {
      const d = await prompt([
        {
          name: 'startDate',
          message: '排程起始日期 (YYYY-MM-DD)：',
          default: new Date().toISOString().slice(0, 10)
        },
        {
          name: 'timeOfDay',
          message: '每天的發佈時間 (24h HH:MM，例如 10:00)：',
          default: '10:00',
          validate: v => /^\d{1,2}:\d{2}$/.test(v) ? true : '格式應為 HH:MM'
        },
        {
          name: 'perDay',
          message: '每天要發佈幾支影片 (整數 > 0)：',
          default: 1,
          validate: v => (Number.isInteger(Number(v)) && Number(v) > 0) ? true : '請輸入正整數'
        },
        {
          name: 'spacingMins',
          message: '同日內多支影片間隔分鐘數（若每天 >1 支，建議至少 1）：',
          default: 1,
          validate: v => (!isNaN(Number(v)) && Number(v) >= 0) ? true : '請輸入數字'
        }
      ]);
      scheduleOptions = {
        type: 'daily',
        startDate: d.startDate,
        timeOfDay: d.timeOfDay,
        perDay: Number(d.perDay),
        spacingMins: Number(d.spacingMins)
      };
    } else {
      scheduleOptions = { type: 'none' };
    }

    const excelPath = path.resolve(basicAnswers.excelPath);
    const sheetName = basicAnswers.sheetName;
    const folderPath = path.resolve(basicAnswers.folderPath);

    const { workbook, worksheet, values } = await readSheetToArrayOfArrays(excelPath, sheetName);
    if (!values || values.length === 0) {
      console.log('該工作表沒有資料。');
      return;
    }

    const headers = values[0].map(h => normalizeHeader(h));
    const idxMap = {};
    ['seo_title_zh', 'seo_title_en', 'description_zh', 'description_en', 'yt_tags', 'hashtags'].forEach(k => {
      const pos = headers.findIndex(h => h === k || h === k.toLowerCase());
      if (pos !== -1) idxMap[k] = pos;
    });

    const idColNormalized = normalizeHeader(basicAnswers.idColumn);
    let idColIndex = headers.findIndex(h => h === idColNormalized);
    if (idColIndex === -1) {
      idColIndex = headers.findIndex(h => h.includes('編號') || h.includes('id'));
    }
    if (idColIndex === -1) {
      console.log(`找不到指定的編號欄位 "${basicAnswers.idColumn}"。現有 header:`);
      console.log(values[0]);
      return;
    }

    let uploadedColIndex = headers.findIndex(h => h === 'uploaded' || h === '已上傳');
    if (uploadedColIndex === -1) {
      uploadedColIndex = values[0].length;
      values[0].push('Uploaded');
    }

    console.log('已讀取 header：', values[0]);
    console.log('將從資料列第2列開始逐列處理（跳過 header）');

    let youtube = null;
    if (basicAnswers.doUpload) {
      const auth = await authorize();
      youtube = google.youtube({ version: 'v3', auth });
    }

    // scheduledCount counts how many uploads (or simulations) we've scheduled so far
    let scheduledCount = 0;

    for (let r = 1; r < values.length; r++) {
      const row = values[r] || [];
      const idCell = row[idColIndex];
      if (!idCell || idCell.toString().trim() === '') {
        console.log(`第 ${r + 1} 列沒有編號，跳過`);
        continue;
      }
      const baseName = idCell.toString().trim();
      const videoPath = findFileByBase(folderPath, baseName);
      if (!videoPath) {
        console.log(`第 ${r + 1} 列編號 ${baseName}：找不到對應檔案於 ${folderPath}，跳過`);
        values[r] = values[r] || [];
        values[r][uploadedColIndex] = `MISSING FILE`;
        continue;
      }

      const titleZH = row[idxMap['seo_title_zh']] || '';
      const titleEN = row[idxMap['seo_title_en']] || '';
      const descZH = row[idxMap['description_zh']] || '';
      const descEN = row[idxMap['description_en']] || '';
      const tagsRaw = row[idxMap['yt_tags']] || '';
      const hashtags = row[idxMap['hashtags']] || '';

      const title = `${titleZH}${titleZH && titleEN ? ' / ' : ''}${titleEN}`.trim() || baseName;
      let description = `${descZH}${descZH && descEN ? '\n\n' : ''}${descEN}`.trim();
      if (hashtags) description += `\n\n${hashtags.toString()}`;
      const tags = tagsRaw ? tagsRaw.toString().split(',').map(t => t.trim()).filter(Boolean) : [];

      // compute publishAt based on scheduledCount (only count items that we actually attempt to upload)
      const publishAt = computePublishAtForIndex(scheduledCount, scheduleOptions);

      console.log(`處理 第 ${r + 1} 列 編號=${baseName} 檔案=${videoPath}`);
      console.log(`  Title: ${title}`);
      console.log(`  publishAt: ${publishAt || '立即公開'}`);

      if (!basicAnswers.doUpload) {
        const timestamp = (new Date()).toISOString();
        values[r] = values[r] || [];
        values[r][uploadedColIndex] = `${timestamp} | SIMULATED | file: ${path.basename(videoPath)} | publishAt: ${publishAt || 'NOW'}`;
        console.log(`  模擬完成（未上傳）`);
        // increment scheduledCount only if scheduling applied (count simulation as scheduled action)
        if (scheduleOptions && scheduleOptions.type !== 'none') scheduledCount++;
        continue;
      }

      // perform upload
      try {
        const res = await youtube.videos.insert({
          part: ['snippet', 'status'],
          requestBody: {
            snippet: { title, description, tags, categoryId: '22' },
            status: (publishAt ? { privacyStatus: 'private', publishAt } : { privacyStatus: 'public' })
          },
          media: { body: fs.createReadStream(videoPath) }
        });
        const videoId = res.data.id;
        console.log(`  上傳成功，videoId=${videoId}`);

        const timestamp = (new Date()).toISOString();
        values[r] = values[r] || [];
        values[r][uploadedColIndex] = `${timestamp} | videoId: ${videoId} | publishAt: ${publishAt || 'NOW'}`;

        // increment scheduledCount after successful scheduling/upload
        if (scheduleOptions && scheduleOptions.type !== 'none') scheduledCount++;
      } catch (err) {
        console.error(`  第 ${r + 1} 列上傳失敗：`, (err && err.errors) ? err.errors : (err && err.message) ? err.message : err);
        values[r] = values[r] || [];
        values[r][uploadedColIndex] = `ERROR: ${(err && err.message) ? err.message : 'upload failed'}`;
        // do not increment scheduledCount on failure (so next video will reuse same scheduled slot)
      }
    }

    await writeArrayOfArraysToSheetAndSave(workbook, worksheet, values, excelPath);
    console.log(`已將上傳狀態寫回 ${excelPath} 的工作表 ${sheetName}（欄位：Uploaded）。`);
    console.log('全部處理完成');
  } catch (err) {
    console.error('發生錯誤：', (err && err.message) ? err.message : err);
  }
}

main();