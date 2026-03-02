const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');
const fs = require('fs');

// 日志文件
const logFile = path.join(require('os').homedir(), 'dbgui_debug.log');
function writeLog(msg) {
  fs.appendFileSync(logFile, new Date().toISOString() + ' ' + JSON.stringify(msg) + '\n');
}

let mainWindow;
let db;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
    }
  });

  mainWindow.loadFile('index.html');
  // mainWindow.webContents.openDevTools();
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// IPC handlers
ipcMain.handle('select-db-file', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: '选择 SQLite 数据库文件',
    filters: [{ name: 'SQLite DB', extensions: ['db', 'sqlite', 'sqlite3'] }],
    properties: ['openFile']
  });
  if (canceled || filePaths.length === 0) {
    return null;
  }
  // open sqlite
  if (db) {
    db.close();
    db = null;
  }
  db = new sqlite3.Database(filePaths[0]);
  return filePaths[0];
});

ipcMain.handle('get-routes', async () => {
  return new Promise((resolve, reject) => {
    if (!db) {
      return resolve([]);
    }
    db.all('SELECT routeId, routeName FROM routes ORDER BY routeName', (err, rows) => {
      if (err) return reject(err);
      resolve(rows);
    });
  });
});

// cycle helpers for 值 column
const BEIJING_OFFSET = 8 * 60 * 60 * 1000; // ms
const ONE_DAY = 24 * 60 * 60 * 1000;
// base day (2025-12-05 00:30 Beijing) cycle number
const BASE_DATEMS = new Date('2025-12-05T00:30:00+08:00').getTime();
const BASE_CYCLE = Math.floor((BASE_DATEMS + BEIJING_OFFSET - 30*60*1000) / ONE_DAY);
const SHIFTS_ORDER = ['白班','夜班','休息1','中班','休息2'];
const BASE_ORDER = ['甲','戊','丁','丙','乙'];

// helper to determine shift from a timestamp
function computeShift(ts) {
  const d = new Date(ts > 1e12 ? ts : ts * 1000);
  let total = d.getHours() * 60 + d.getMinutes();
  total -= 30;
  if (total < 0) total += 24 * 60;
  if (total < 8 * 60) return '夜班';
  if (total < 16 * 60) return '白班';
  return '中班';
}


function getCycleNumber(ms) {
  const adj = ms + BEIJING_OFFSET - 30*60*1000;
  return Math.floor(adj / ONE_DAY);
}

function rotateArray(arr, n) {
  const len = arr.length;
  const k = ((n % len) + len) % len;
  return arr.slice(len - k).concat(arr.slice(0, len - k));
}

function computeValueFor(ts, shiftName) {
  const ms = ts > 1e12 ? ts : ts * 1000;
  const cycle = getCycleNumber(ms);
  const delta = cycle - BASE_CYCLE;
  const order = rotateArray(BASE_ORDER, delta);
  const idx = SHIFTS_ORDER.indexOf(shiftName);
  if (idx === -1) return '';
  return order[idx];
}

ipcMain.handle('query-records', async (event, { routeId, date }) => {
  return new Promise((resolve, reject) => {
    if (!db) return resolve([]);

    function toBeijing(dateString) {
      return new Date(dateString + '+08:00');
    }
    const start = toBeijing(date + 'T00:30:00');
    const end = new Date(start);
    end.setDate(end.getDate() + 1);
    const startTsSec = Math.floor(start.getTime() / 1000);
    const endTsSec = Math.floor(end.getTime() / 1000);
    const startTsMs = start.getTime();
    const endTsMs = end.getTime();

    // logging intentionally minimal now
    writeLog({step: 'query-records START', routeId, date, startTsSec, endTsSec});

    // Step 1: Get all points for this route
    const sql1 = `
      SELECT DISTINCT p.pointId, p.name as pointName
      FROM points p
      WHERE p.routeId = ?
    `;

    // Step 2: Get route name
    const sql2 = `
      SELECT routeName FROM routes WHERE routeId = ?
    `;

    // Step 3: Get min freqHours for each point
    const sql3 = `
      SELECT pointId, MIN(freqHours) as minFreq
      FROM check_items
      WHERE equipmentId IN (
        SELECT equipmentId FROM equipments WHERE pointId = ?
      )
      GROUP BY pointId
    `;

    // Step 4: Get actual inspection records for timestamp checking
    // include record/session identifiers and operator info for logging
    const sql4 = `
      SELECT rec.recordId,
             rec.sessionId,
             s.operatorId,
             e.employeeName,
             s.shiftId,
             sh.name as shiftName,
             rec.pointId,
             rec.timestamp
      FROM inspection_records rec
      JOIN inspection_sessions s ON rec.sessionId = s.sessionId
      LEFT JOIN shifts sh ON s.shiftId = sh.shiftId
      LEFT JOIN employees e ON s.operatorId = e.employeeId
      WHERE s.routeId = ? AND rec.timestamp >= ? AND rec.timestamp < ?
      ORDER BY rec.timestamp
    `;

    // Step 5: Get employee info for the date (to match shifts with operators)
    const sql5 = `
      SELECT DISTINCT s.shiftId, sh.name as shiftName, e.employeeId, e.employeeName
      FROM inspection_sessions s
      LEFT JOIN employees e ON s.operatorId = e.employeeId
      LEFT JOIN shifts sh ON s.shiftId = sh.shiftId
      WHERE s.routeId = ? AND s.startTime >= ? AND s.startTime < ?
    `;

    function computeShift(ts) {
      const d = new Date(ts > 1e12 ? ts : ts * 1000);
      let total = d.getHours() * 60 + d.getMinutes();
      total -= 30;
      if (total < 0) total += 24 * 60;
      if (total < 8 * 60) return '夜班';
      if (total < 16 * 60) return '白班';
      return '中班';
    }

    function processResults(points, allPoints, minFreqMap, records, shifts, routeName) {
      const result = [];
      const options = { timeZone: 'Asia/Shanghai', hour12: false };
      
      // Group actual records by shift+pointId to check slot existence
      // include employeeName so we can assign later
      const recordsByShiftPoint = {};
      records.forEach(r => {
        const shiftName = r.shiftName || computeShift(r.timestamp);
        const key = shiftName + '|' + r.pointId;
        if (!recordsByShiftPoint[key]) recordsByShiftPoint[key] = [];
        recordsByShiftPoint[key].push({
          timestamp: r.timestamp,
          shift: shiftName,
          pointId: r.pointId,
          employeeName: r.employeeName || ''
        });
      });

      // Create a map of shift names to employee names
      const shiftEmployeeMap = {};
      shifts.forEach(s => {
        const key = s.shiftName;
        if (!shiftEmployeeMap[key]) shiftEmployeeMap[key] = s.employeeName || '';
      });
      writeLog({step:'shiftEmployeeMap', map: shiftEmployeeMap});

      // Get Beijing date for display
      const beijingDate = start.toLocaleDateString('zh-CN', options);

      // Define shifts for this day
      const shiftsForDay = ['夜班', '白班', '中班'];

      // For each point
      points.forEach(point => {
        const minFreq = minFreqMap[point.pointId] || 4;
        const expectedSlots = Math.max(1, Math.floor(8 / minFreq));

        // Calculate shift boundaries
        const shiftConfigs = {
          '夜班': {startHour: 0, startMin: 30, durationHours: 8},
          '白班': {startHour: 8, startMin: 30, durationHours: 8},
          '中班': {startHour: 16, startMin: 30, durationHours: 8}
        };

        // For each shift
        shiftsForDay.forEach(shiftName => {
          const config = shiftConfigs[shiftName];
          let shiftStart = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 
                                   config.startHour, config.startMin);
          
          const shiftStartTs = shiftStart.getTime() / 1000;
          const slotDuration = config.durationHours / expectedSlots; // hours per slot

          const key = shiftName + '|' + point.pointId;
          const recordsInSlot = recordsByShiftPoint[key] || [];
          // prefer employee name from first record if available
          let employeeName = '';
          if (recordsInSlot.length > 0) {
            employeeName = recordsInSlot[0].employeeName || '';
          }

          // Generate one row per slot for this point in this shift
          for (let slot = 1; slot <= expectedSlots; slot++) {
            const slotStartSec = shiftStartTs + (slot - 1) * slotDuration * 3600;
            const slotEndSec = slotStartSec + slotDuration * 3600;

            // Check if there's any record in this slot
            const hasRecord = recordsInSlot.some(recordObj => {
              const ts = recordObj.timestamp;
              const tsNum = ts > 1e12 ? ts / 1000 : ts;
              const inRange = tsNum >= slotStartSec && tsNum < slotEndSec;
              if (point.pointId && slot === 1 && recordsInSlot.length > 0) {
                writeLog({
                  step: 'slot check',
                  shift: shiftName,
                  point: point.pointId,
                  slot,
                  slotStart: slotStartSec,
                  slotEnd: slotEndSec,
                  recordTs: tsNum,
                  inRange
                });
              }
              return inRange;
            });

            const valueForShift = computeValueFor(start.getTime(), shiftName);

            const rowObj = {
              date: beijingDate,
              time: '',
              routeName: routeName,
              employeeName: employeeName,
              calcValue: valueForShift,
              shiftName: shiftName,
              pointId: point.pointId,
              pointName: point.pointName,
              freqHours: minFreq,
              slotIndex: slot,
              inspectionStatus: hasRecord ? '已检' : '未检'
            };
          writeLog({step:'generatedRow', row: rowObj});
          result.push(rowObj);
          }
        });
      });

      return result;
    }

    // Execute queries
    db.all(sql1, [routeId], (err, points) => {
      if (err) return reject(err);
      writeLog({step: 'sql1 returned', pointCount: points ? points.length : 0});
      if (points.length === 0) return resolve([]);

      db.get(sql2, [routeId], (err, routeRow) => {
        if (err) return reject(err);
        const routeName = routeRow ? routeRow.routeName : '';

        // Get unique pointIds
        const pointIds = points.map(p => p.pointId);
        
        // Get minFreq for each point
        const minFreqMap = {};
        let completed = 0;

        pointIds.forEach(pointId => {
          db.get(sql3, [pointId], (err, row) => {
            if (row) minFreqMap[pointId] = row.minFreq;
            completed++;
            
            if (completed === pointIds.length) {
              // All min freqs fetched, now get actual records and shift info
              writeLog({step: 'sql4 query START', routeId, startTsSec, endTsSec});
              // first attempt with seconds-based timestamps
              db.all(sql4, [routeId, startTsSec, endTsSec], (err, records) => {
                if (err) {
                  writeLog({step: 'sql4 ERROR', err: err.message});
                  return reject(err);
                }
                // if no rows, retry using millisecond range
                if (!records || records.length === 0) {
                  writeLog({step: 'sql4 returned 0 with seconds, retrying with ms'});
                  db.all(sql4, [routeId, startTsMs, endTsMs], (err2, records2) => {
                    if (err2) {
                      writeLog({step: 'sql4 ERROR ms', err: err2.message});
                      return reject(err2);
                    }
                    records = records2 || [];
                    afterRecords();
                  });
                } else {
                  afterRecords();
                }

                function afterRecords() {
                  writeLog({step: 'sql4 returned', recordCount: records ? records.length : 0});
                  if (records && records.length > 0) {
                    writeLog({step: 'Sample records', samples: records.slice(0, 2)});
                  }
                  // no console logging here anymore

                  // log each record by shift order with detailed info
                  if (records && records.length > 0) {
                    const shiftOrder = ['夜班', '白班', '中班'];
                    records.sort((a, b) => shiftOrder.indexOf(a.shiftName) - shiftOrder.indexOf(b.shiftName));
                    records.forEach(r => {
                      // compute slot index for the record timestamp
                      const minFreq = minFreqMap[r.pointId] || 4;
                      const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
                      const shiftConfigs = {
                        '夜班': {startHour: 0, startMin: 30, durationHours: 8},
                        '白班': {startHour: 8, startMin: 30, durationHours: 8},
                        '中班': {startHour: 16, startMin: 30, durationHours: 8}
                      };
                      const cfg = shiftConfigs[r.shiftName];
                      let shiftStart = new Date(start.getFullYear(), start.getMonth(), start.getDate(), cfg.startHour, cfg.startMin);
                      if (r.shiftName === '中班') {
                        // middle shift extends into next day
                        shiftStart = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 1, cfg.startHour, cfg.startMin);
                      }
                      const shiftStartTs = shiftStart.getTime() / 1000;
                      const slotDuration = cfg.durationHours / expectedSlots;
                      const tsNum = r.timestamp > 1e12 ? r.timestamp / 1000 : r.timestamp;
                      let slotIndex = 0;
                      for (let i = 1; i <= expectedSlots; i++) {
                        const sstart = shiftStartTs + (i - 1) * slotDuration * 3600;
                        const send = sstart + slotDuration * 3600;
                        if (tsNum >= sstart && tsNum < send) {
                          slotIndex = i;
                          break;
                        }
                      }
                      const timeStr = new Date(tsNum * 1000).toLocaleString('zh-CN', {timeZone: 'Asia/Shanghai'});
                      writeLog({detailLog:`record ${r.recordId} session ${r.sessionId} point ${r.pointId} slot ${slotIndex} ts ${timeStr} operator ${r.operatorId}(${r.employeeName || ''})`});
                    });
                  }

                  db.all(sql5, [routeId, startTsSec, endTsSec], (err, shifts) => {
                    if (err) return reject(err);

                    const finalResult = processResults(points, pointIds, minFreqMap, records || [], shifts || [], routeName);
                    writeLog({step: 'finalResult', count: finalResult.length});
                    resolve(finalResult);
                  });
                }
              });
            }
          });
        });
      });
    });
  });
});

// handler for detail lookup
ipcMain.handle('get-record-details', async (event, { routeId, date, shiftName, pointId, slotIndex }) => {
  writeLog({step:'get-details', routeId, date, shiftName, pointId, slotIndex});
  // compute timestamp boundaries for the given slot
  function toBeijing(dateString) { return new Date(dateString + '+08:00'); }
  const startDay = toBeijing(date + 'T00:30:00');
  const d = new Date(startDay);
  const shiftConfigs = {
    '夜班': {startHour:0, startMin:30},
    '白班': {startHour:8, startMin:30},
    '中班': {startHour:16, startMin:30}
  };
  let shiftStart = new Date(d.getFullYear(), d.getMonth(), d.getDate(), shiftConfigs[shiftName].startHour, shiftConfigs[shiftName].startMin);
  if (shiftName === '中班') {
    // middle shift spills into next day
    shiftStart = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1, shiftConfigs[shiftName].startHour, shiftConfigs[shiftName].startMin);
  }
  // fetch minFreq for point
  const minFreqRow = await new Promise((res,rej)=>{
    db.get(`SELECT MIN(freqHours) as minFreq FROM check_items WHERE equipmentId IN (SELECT equipmentId FROM equipments WHERE pointId = ?)`, [pointId], (e,r)=>{ if(e) rej(e); else res(r); });
  });
  const minFreq = (minFreqRow && minFreqRow.minFreq) || 4;
  const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
  const slotDuration = 8 / expectedSlots; // hours
  const slotStartSec = shiftStart.getTime()/1000 + (slotIndex-1)*slotDuration*3600;
  const slotEndSec = slotStartSec + slotDuration*3600;

  const sql = `
    SELECT rec.recordId,
           rec.timestamp,
           p.pointId,
           p.name as pointName,
           eq.equipmentName,
           ci.itemId,
           ci.itemName as checkItemName,
           ci.freqHours,
           ri.slotIndex,
           ri.value,
           ri.abnormal
    FROM inspection_records rec
    JOIN inspection_record_items ri ON rec.recordId = ri.recordId
    LEFT JOIN points p ON rec.pointId = p.pointId
    LEFT JOIN equipments eq ON ri.equipmentId = eq.equipmentId
    LEFT JOIN check_items ci ON ri.itemId = ci.itemId
    JOIN inspection_sessions s ON rec.sessionId = s.sessionId
    WHERE s.routeId = ?
      AND rec.pointId = ?
      AND rec.timestamp >= ?
      AND rec.timestamp < ?
    ORDER BY rec.timestamp
  `;

  return new Promise((resolve, reject) => {
    writeLog({step:'get-details sql', sql, params:[routeId, pointId, slotStartSec, slotEndSec]});
    // first try seconds
    db.all(sql, [routeId, pointId, slotStartSec, slotEndSec], (err, rows) => {
      if (err) return reject(err);
      if (!rows || rows.length === 0) {
        // attempt ms range
        const slotStartMs = slotStartSec * 1000;
        const slotEndMs = slotEndSec * 1000;
        writeLog({step:'get-details retry ms', params:[routeId, pointId, slotStartMs, slotEndMs]});
        db.all(sql, [routeId, pointId, slotStartMs, slotEndMs], (err2, rows2) => {
          if (err2) return reject(err2);
          writeLog({step:'get-details returned', count: rows2 ? rows2.length : 0});
          resolve((rows2||[]).map(r => {
            const d2 = new Date(r.timestamp > 1e12 ? r.timestamp : r.timestamp * 1000);
            const options = { timeZone: 'Asia/Shanghai', hour12: false };
            return {
              date: d2.toLocaleDateString('zh-CN', options),
              time: d2.toLocaleTimeString('zh-CN', options),
              pointId: r.pointId,
              pointName: r.pointName,
              equipmentName: r.equipmentName,
              itemId: r.itemId,
              itemName: r.checkItemName,
              freqHours: r.freqHours,
              slotIndex: r.slotIndex,
              value: r.value,
              abnormal: r.abnormal
            };
          }));
        });
      } else {
        writeLog({step:'get-details returned', count: rows ? rows.length : 0});
        resolve(rows.map(r => {
          const d2 = new Date(r.timestamp > 1e12 ? r.timestamp : r.timestamp * 1000);
          const options = { timeZone: 'Asia/Shanghai', hour12: false };
          return {
            date: d2.toLocaleDateString('zh-CN', options),
            time: d2.toLocaleTimeString('zh-CN', options),
            pointId: r.pointId,
            pointName: r.pointName,
            equipmentName: r.equipmentName,
            itemId: r.itemId,
            itemName: r.checkItemName,
            freqHours: r.freqHours,
            slotIndex: r.slotIndex,
            value: r.value,
            abnormal: r.abnormal
          };
        }));
      }
    });
  });
});

ipcMain.handle('export-records', async (event, params) => {
  // params can be either {records, outputPath} (legacy) or
  // {routeId, date, outputPath} for exporting full-details.
  let rows = [];

  if (params.routeId && params.date) {
    // generate detailed rows for the given day and route
    const {routeId, date} = params;

    // helper from queryRecords
    function toBeijing(dateString) { return new Date(dateString + '+08:00'); }
    const start = toBeijing(date + 'T00:30:00');
    const end = new Date(start);
    end.setDate(end.getDate() + 1);
    const startTsSec = Math.floor(start.getTime() / 1000);
    const endTsSec = Math.floor(end.getTime() / 1000);
    const startTsMs = start.getTime();
    const endTsMs = end.getTime();

    // we will need routeName, points list, minFreq map
    const points = await new Promise((res, rej) => {
      db.all(`SELECT DISTINCT p.pointId, p.name as pointName
              FROM points p
              WHERE p.routeId = ?`, [routeId], (err, rows) => {
        if (err) return rej(err);
        res(rows || []);
      });
    });
    const routeName = await new Promise((res, rej) => {
      db.get(`SELECT routeName FROM routes WHERE routeId = ?`, [routeId], (err, row) => {
        if (err) return rej(err);
        res(row ? row.routeName : '');
      });
    });

    const minFreqMap = {};
    await Promise.all(points.map(pt => {
      return new Promise((res, rej) => {
        db.get(`SELECT MIN(freqHours) as minFreq
                FROM check_items
                WHERE equipmentId IN (
                  SELECT equipmentId FROM equipments WHERE pointId = ?
                )`, [pt.pointId], (err, row) => {
          if (!err && row) minFreqMap[pt.pointId] = row.minFreq;
          res();
        });
      });
    }));

    // fetch all item-level records for the day
    const itemSql = `
      SELECT rec.recordId,
             rec.timestamp,
             sh.name as shiftName,
             e.employeeName,
             rec.pointId,
             p.name as pointName,
             eq.equipmentName,
             ci.itemName as checkItemName,
             ci.freqHours,
             ri.slotIndex,
             ri.value,
             ri.abnormal
      FROM inspection_records rec
      JOIN inspection_record_items ri ON rec.recordId = ri.recordId
      LEFT JOIN inspection_sessions s ON rec.sessionId = s.sessionId
      LEFT JOIN shifts sh ON s.shiftId = sh.shiftId
      LEFT JOIN employees e ON s.operatorId = e.employeeId
      LEFT JOIN points p ON rec.pointId = p.pointId
      LEFT JOIN equipments eq ON ri.equipmentId = eq.equipmentId
      LEFT JOIN check_items ci ON ri.itemId = ci.itemId
      WHERE s.routeId = ?
        AND rec.timestamp >= ?
        AND rec.timestamp < ?
      ORDER BY rec.timestamp
    `;

    let items = [];
    await new Promise((res, rej) => {
      db.all(itemSql, [routeId, startTsSec, endTsSec], (err, rows) => {
        if (err) return rej(err);
        if (!rows || rows.length === 0) {
          // try ms
          db.all(itemSql, [routeId, startTsMs, endTsMs], (err2, rows2) => {
            if (err2) return rej(err2);
            items = rows2 || [];
            res();
          });
        } else {
          items = rows;
          res();
        }
      });
    });

    // map each item to a row and keep track of filled slots
    const filledSlots = {};
    const options = { timeZone: 'Asia/Shanghai', hour12: false };

    function computeSlotIndex(ts, shiftName, pointId) {
      const minFreq = minFreqMap[pointId] || 4;
      const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
      const shiftConfigs = {
        '夜班': {startHour: 0, startMin: 30, durationHours: 8},
        '白班': {startHour: 8, startMin: 30, durationHours: 8},
        '中班': {startHour: 16, startMin: 30, durationHours: 8}
      };
      let shiftStart = new Date(start.getFullYear(), start.getMonth(), start.getDate(),
                               shiftConfigs[shiftName].startHour,
                               shiftConfigs[shiftName].startMin);
      if (shiftName === '中班') {
        shiftStart = new Date(start.getFullYear(), start.getMonth(), start.getDate()+1,
                               shiftConfigs[shiftName].startHour,
                               shiftConfigs[shiftName].startMin);
      }
      const shiftStartTs = shiftStart.getTime() / 1000;
      const slotDuration = shiftConfigs[shiftName].durationHours / expectedSlots;
      const tsNum = ts > 1e12 ? ts / 1000 : ts;
      for (let i = 1; i <= expectedSlots; i++) {
        const sstart = shiftStartTs + (i - 1) * slotDuration * 3600;
        const send = sstart + slotDuration * 3600;
        if (tsNum >= sstart && tsNum < send) return i;
      }
      return 0;
    }

    items.forEach(r => {
      const d2 = new Date(r.timestamp > 1e12 ? r.timestamp : r.timestamp * 1000);
      const dateStr = d2.toLocaleDateString('zh-CN', options);
      const timeStr = d2.toLocaleTimeString('zh-CN', options);
      const slotIdx = computeSlotIndex(r.timestamp, r.shiftName || '', r.pointId);
      const calcValue = computeValueFor(start.getTime(), r.shiftName || '');
      const abnormalText = r.abnormal === 0 ? '正常' : (r.abnormal === 1 ? '异常' : (r.abnormal || ''));

      rows.push({
        date: dateStr,
        time: timeStr,
        routeName,
        employeeName: r.employeeName || '',
        calcValue,
        shiftName: r.shiftName || '',
        pointId: r.pointId,
        pointName: r.pointName,
        equipmentName: r.equipmentName || '',
        checkItemName: r.checkItemName || '',
        freqHours: r.freqHours,
        slotIndex: slotIdx,
        value: r.value,
        abnormal: abnormalText
      });

      const key = `${r.shiftName}|${r.pointId}|${slotIdx}`;
      filledSlots[key] = true;
    });

    // now add uninspected placeholder rows
    const beijingDate = start.toLocaleDateString('zh-CN', options);
    const shiftsForDay = ['夜班', '白班', '中班'];
    points.forEach(point => {
      const minFreq = minFreqMap[point.pointId] || 4;
      const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
      shiftsForDay.forEach(shiftName => {
        for (let slot = 1; slot <= expectedSlots; slot++) {
          const key = `${shiftName}|${point.pointId}|${slot}`;
          if (!filledSlots[key]) {
            rows.push({
              date: beijingDate,
              time: '',
              routeName,
              employeeName: '',
              calcValue: computeValueFor(start.getTime(), shiftName),
              shiftName,
              pointId: point.pointId,
              pointName: point.pointName,
              equipmentName: '',
              checkItemName: '',
              freqHours: minFreq,
              slotIndex: slot,
              value: '',
              abnormal: '未检'
            });
          }
        }
      });
    });

  } else if (params.records && params.outputPath) {
    // legacy behaviour: simply dump passed records
    rows = params.records;
  }

  // write the workbook
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('点检详情');
  sheet.columns = [
    {header:'日期', key:'date'},
    {header:'时间', key:'time'},
    {header:'路线名', key:'routeName'},
    {header:'点检人', key:'employeeName'},
    {header:'值', key:'calcValue'},
    {header:'班次', key:'shiftName'},
    {header:'点位ID', key:'pointId'},
    {header:'点位名', key:'pointName'},
    {header:'设备名', key:'equipmentName'},
    {header:'点检项', key:'checkItemName'},
    {header:'点检频率', key:'freqHours'},
    {header:'点检频次', key:'slotIndex'},
    {header:'检测值', key:'value'},
    {header:'是否正常', key:'abnormal'}
  ];
  sheet.columns = sheet.columns;

  // sort rows for more predictable output
  const shiftOrder = ['夜班','白班','中班'];
  rows.sort((a,b) => {
    const s = shiftOrder.indexOf(a.shiftName) - shiftOrder.indexOf(b.shiftName);
    if (s !== 0) return s;
    if (a.pointId && b.pointId) {
      const cmp = a.pointId.localeCompare(b.pointId);
      if (cmp !== 0) return cmp;
    }
    return (a.slotIndex || 0) - (b.slotIndex || 0);
  });

  rows.forEach(r => {
    const row = sheet.addRow(r);
    if (r.abnormal === '未检') {
      const abnormalCell = row.getCell('abnormal');
      abnormalCell.font = { color: { argb: 'FFFF0000' } };
    }
  });

  await workbook.xlsx.writeFile(params.outputPath);
  return true;
});

// helper to show save dialog
ipcMain.handle('select-export-path', async () => {
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: '保存导出结果',
    defaultPath: 'inspection_results.xlsx',
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });
  if (canceled) return { canceled: true };
  return { canceled: false, filePath };
});

// Get timestamp range from inspection_records
ipcMain.handle('get-timestamp-range', async () => {
  return new Promise((resolve, reject) => {
    if (!db) return resolve({ minTimestamp: null, maxTimestamp: null });
    
    const sql = `
      SELECT MIN(timestamp) as minTs, MAX(timestamp) as maxTs
      FROM inspection_records
    `;
    
    db.get(sql, (err, row) => {
      if (err) return reject(err);
      if (!row || !row.minTs || !row.maxTs) {
        return resolve({ minTimestamp: null, maxTimestamp: null });
      }
      // Convert timestamp to seconds if needed for consistency
      const minTs = row.minTs > 1e12 ? Math.floor(row.minTs / 1000) : row.minTs;
      const maxTs = row.maxTs > 1e12 ? Math.floor(row.maxTs / 1000) : row.maxTs;
      resolve({ minTimestamp: minTs, maxTimestamp: maxTs });
    });
  });
});

// Query miss rate statistics
ipcMain.handle('query-miss-rate-stats', async (event, { routeId, startDate, endDate }) => {
  return new Promise((resolve, reject) => {
    if (!db) return resolve({ leftData: [], rightData: [] });

    function toBeijing(dateString) {
      return new Date(dateString + '+08:00');
    }
    
    const startDayStart = toBeijing(startDate + 'T00:30:00');
    const endDayStart = toBeijing(endDate + 'T00:30:00');
    const endDayEnd = new Date(endDayStart);
    endDayEnd.setDate(endDayEnd.getDate() + 1);
    
    const startTsSec = Math.floor(startDayStart.getTime() / 1000);
    const endTsSec = Math.floor(endDayEnd.getTime() / 1000);
    const startTsMs = startDayStart.getTime();
    const endTsMs = endDayEnd.getTime();

    // Get all points for this route
    const sqlPoints = `
      SELECT DISTINCT p.pointId, p.name as pointName
      FROM points p
      WHERE p.routeId = ?
    `;

    // Get min freq for each point
    const sqlMinFreq = `
      SELECT MIN(freqHours) as minFreq
      FROM check_items
      WHERE equipmentId IN (
        SELECT equipmentId FROM equipments WHERE pointId = ?
      )
    `;

    // Get actual inspection records in date range
    const sqlRecords = `
      SELECT rec.recordId,
             rec.sessionId,
             rec.pointId,
             s.shiftId,
             sh.name as shiftName,
             rec.timestamp
      FROM inspection_records rec
      JOIN inspection_sessions s ON rec.sessionId = s.sessionId
      LEFT JOIN shifts sh ON s.shiftId = sh.shiftId
      WHERE s.routeId = ? AND rec.timestamp >= ? AND rec.timestamp < ?
      ORDER BY rec.timestamp
    `;

    // Helper function to compute shift from timestamp (fallback)
    function computeShift(ts) {
      const d = new Date(ts > 1e12 ? ts : ts * 1000);
      let total = d.getHours() * 60 + d.getMinutes();
      total -= 30;
      if (total < 0) total += 24 * 60;
      if (total < 8 * 60) return '夜班';
      if (total < 16 * 60) return '白班';
      return '中班';
    }

    db.all(sqlPoints, [routeId], (err, points) => {
      if (err) return reject(err);
      if (!points || points.length === 0) {
        return resolve({ leftData: [], rightData: [] });
      }

      // Get minFreq for each point
      const minFreqMap = {};
      let freqCompleted = 0;

      points.forEach(point => {
        db.get(sqlMinFreq, [point.pointId], (err, row) => {
          if (row) minFreqMap[point.pointId] = row.minFreq;
          freqCompleted++;

          if (freqCompleted === points.length) {
            // All minFreqs fetched, now get inspection records
            db.all(sqlRecords, [routeId, startTsSec, endTsSec], (err, records) => {
              writeLog({step:'sqlRecords returned', count: records ? records.length : 0, err: err ? err.message : null});
              if (records && records.length > 0) {
                writeLog({step:'sqlRecords sample', samples: records.slice(0,5)});
              }
              if (err) {
                // Try with ms timestamps
                db.all(sqlRecords, [routeId, startTsMs, endTsMs], (err2, records2) => {
                  writeLog({step:'sqlRecords ms returned', count: records2 ? records2.length : 0, err: err2 ? err2.message : null});
                  if (records2 && records2.length > 0) {
                    writeLog({step:'sqlRecords ms sample', samples: records2.slice(0,5)});
                  }
                  if (err2) return reject(err2);
                  processStats(points, minFreqMap, records2 || []);
                });
              } else {
                if (!records || records.length === 0) {
                  // Try with ms timestamps
                  db.all(sqlRecords, [routeId, startTsMs, endTsMs], (err2, records2) => {
                    writeLog({step:'sqlRecords ms returned (fallback)', count: records2 ? records2.length : 0, err: err2 ? err2.message : null});
                    if (records2 && records2.length > 0) {
                      writeLog({step:'sqlRecords ms sample (fallback)', samples: records2.slice(0,5)});
                    }
                    if (err2) return reject(err2);
                    processStats(points, minFreqMap, records2 || []);
                  });
                } else {
                  processStats(points, minFreqMap, records || []);
                }
              }
            });
          }
        });
      });

      function processStats(points, minFreqMap, records) {
        const options = { timeZone: 'Asia/Shanghai', hour12: false };
        
        // Log received records for debugging
        writeLog({step: 'processStats START', recordCount: records.length, firstFewRecords: records.slice(0, 3).map(r => ({shiftName: r.shiftName, timestamp: r.timestamp, pointId: r.pointId}))});
        
        // Build a map of which slots were inspected
        // Key: YYYY-MM-DD|shiftName|pointId|slotIndex
        const inspectedSlots = {};
        records.forEach(r => {
          const ts = r.timestamp > 1e12 ? r.timestamp / 1000 : r.timestamp;
          const d = new Date(ts * 1000);
          
          // compute shift name based solely on timestamp (DB values may be wrong)
          const computed = computeShift(ts);
          if (r.shiftName && r.shiftName !== computed) {
            writeLog({step:'shiftName mismatch', recordId: r.recordId, dbShift: r.shiftName, computed});
          }
          const shiftName = computed;
          const pointId = r.pointId;
          
          // Calculate slot index
          const minFreq = minFreqMap[pointId] || 4;
          const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
          
          const shiftConfigs = {
            '夜班': {startHour: 0, startMin: 30},
            '白班': {startHour: 8, startMin: 30},
            '中班': {startHour: 16, startMin: 30}
          };
          
          let shiftStart;
          if (shiftName === '中班') {
            // middle shift covers 16:30–00:30; determine correct start day
            shiftStart = new Date(d.getFullYear(), d.getMonth(), d.getDate(),
                                   shiftConfigs['中班'].startHour,
                                   shiftConfigs['中班'].startMin);
            // if timestamp is before 8:00 (early morning) it belongs to previous day's 中班
            if (d.getHours() < 8) {
              shiftStart.setDate(shiftStart.getDate() - 1);
            }
          } else {
            shiftStart = new Date(d.getFullYear(), d.getMonth(), d.getDate(),
                                   shiftConfigs[shiftName]?.startHour || 0,
                                   shiftConfigs[shiftName]?.startMin || 30);
          }
          
          const shiftStartTs = shiftStart.getTime() / 1000;
          const slotDuration = 8 / expectedSlots; // hours
          
          let slotIndex = 0;
          for (let i = 1; i <= expectedSlots; i++) {
            const slotStart = shiftStartTs + (i - 1) * slotDuration * 3600;
            const slotEnd = slotStart + slotDuration * 3600;
            if (ts >= slotStart && ts < slotEnd) {
              slotIndex = i;
              break;
            }
          }
          
          // determine date string for the shift; for 中班, use the day when the shift started
          let dateStr;
          if (shiftName === '中班') {
            // for timestamps after midnight (hour < 8) we already adjusted shiftStart below
            const base = new Date(shiftStart);
            if (shiftStart.getHours() === 16) {
              // shiftStart is correct day for 中班; subtract one if timestamp was early morning
              if (d.getHours() < 8) {
                base.setDate(base.getDate() - 1);
              }
            }
            dateStr = base.toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
          } else {
            dateStr = d.toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
          }
          
          // Only record if slotIndex was found (not 0)
          if (slotIndex > 0) {
            const key = `${dateStr}|${shiftName}|${pointId}|${slotIndex}`;
            inspectedSlots[key] = true;
          } else {
            writeLog({step: 'WARNING: slotIndex not found', dateStr, shiftName, pointId, ts, shiftStartTs, expectedSlots});
          }
        });

        // Left data: For each date+shift combination, calculate miss rate
        const leftDataMap = {};
        
        // Loop through all dates in range (inclusive)
        // Note: Query range is [startDayStart, endDayEnd), so we need to process until endDayEnd
        let currentDate = new Date(startDayStart);
        while (currentDate < endDayEnd) {
          const dateStr = currentDate.toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
          const shifts = ['夜班', '白班', '中班'];
          
          shifts.forEach(shiftName => {
            let inspectedCount = 0;
            let uninspectedCount = 0;
            
            points.forEach(point => {
              const minFreq = minFreqMap[point.pointId] || 4;
              const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
              
              for (let slot = 1; slot <= expectedSlots; slot++) {
                const key = `${dateStr}|${shiftName}|${point.pointId}|${slot}`;
                if (inspectedSlots[key]) {
                  inspectedCount++;
                } else {
                  uninspectedCount++;
                }
              }
            });
            
            if (inspectedCount + uninspectedCount > 0) {
              const missRate = ((uninspectedCount / (inspectedCount + uninspectedCount)) * 100).toFixed(1);
              const valueForDay = computeValueFor(currentDate.getTime(), shiftName);
              
              const key = `${dateStr}|${shiftName}`;
              leftDataMap[key] = {
                date: dateStr,
                value: valueForDay,
                shift: shiftName,
                inspected: inspectedCount,
                uninspected: uninspectedCount,
                missRate: missRate + '%'
              };
            }
          });
          
          currentDate.setDate(currentDate.getDate() + 1);
        }
        
        // Convert map to array and sort
        const leftData = Object.values(leftDataMap).sort((a, b) => {
          // Parse date string in format YYYY/MM/DD to comparable format
          const parseDate = (dateStr) => {
            const parts = dateStr.split('/');
            return new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2])).getTime();
          };
          const timeA = parseDate(a.date);
          const timeB = parseDate(b.date);
          if (timeA !== timeB) {
            return timeA - timeB;
          }
          const shiftOrder = ['夜班', '白班', '中班'];
          return shiftOrder.indexOf(a.shift) - shiftOrder.indexOf(b.shift);
        });
        
        // Right data: For each value (甲乙丙丁戊) in date range, calculate total miss rate
        const rightDataMap = {};
        const values = ['甲', '乙', '丙', '丁', '戊'];
        
        values.forEach(value => {
          let totalInspected = 0;
          let totalUninspected = 0;
          
          currentDate = new Date(startDayStart);
          while (currentDate < endDayEnd) {
            const dateStr = currentDate.toLocaleDateString('zh-CN', { timeZone: 'Asia/Shanghai' });
            const shifts = ['夜班', '白班', '中班'];
            
            shifts.forEach(shiftName => {
              // Check if this shift+date has this value
              const shiftValue = computeValueFor(currentDate.getTime(), shiftName);
              if (shiftValue === value) {
                points.forEach(point => {
                  const minFreq = minFreqMap[point.pointId] || 4;
                  const expectedSlots = Math.max(1, Math.floor(8 / minFreq));
                  
                  for (let slot = 1; slot <= expectedSlots; slot++) {
                    const key = `${dateStr}|${shiftName}|${point.pointId}|${slot}`;
                    if (inspectedSlots[key]) {
                      totalInspected++;
                    } else {
                      totalUninspected++;
                    }
                  }
                });
              }
            });
            
            currentDate.setDate(currentDate.getDate() + 1);
          }
          
          if (totalInspected + totalUninspected > 0) {
            const missRate = ((totalUninspected / (totalInspected + totalUninspected)) * 100).toFixed(1);
            rightDataMap[value] = {
              value,
              inspected: totalInspected,
              uninspected: totalUninspected,
              missRate: missRate + '%'
            };
          }
        });
        
        const rightData = Object.values(rightDataMap);
        
        resolve({ leftData, rightData });
      }
    });
  });
});
