const btnSelectDb = document.getElementById('btn-select-db');
const routeSelect = document.getElementById('route-select');
const dateInput = document.getElementById('date-input');
const btnPreview = document.getElementById('btn-preview');
const btnExport = document.getElementById('btn-export');
const tbody = document.querySelector('#results-table tbody');

// Miss rate page elements
const btnSelectDbMissrate = document.getElementById('btn-select-db-missrate');
const routeSelectMissrate = document.getElementById('route-select-missrate');
const startDateInput = document.getElementById('start-date-input');
const endDateInput = document.getElementById('end-date-input');
const btnQueryMissrate = document.getElementById('btn-query-missrate');
const missrateLeftTbody = document.querySelector('#missrate-left-table tbody');
const missrateRightTbody = document.querySelector('#missrate-right-table tbody');

let currentRecords = [];
let currentPage = 'detail'; // 'detail' or 'missrate'
let timestampRange = { minTimestamp: null, maxTimestamp: null };

// Page switching
function showPage(page) {
  currentPage = page;
  const detailPage = document.getElementById('detail-page');
  const missratePage = document.getElementById('missrate-page');
  
  if (page === 'detail') {
    detailPage.style.display = 'flex';
    missratePage.style.display = 'none';
  } else {
    detailPage.style.display = 'none';
    missratePage.style.display = 'flex';
  }
}

// Detail page: Select database
btnSelectDb.addEventListener('click', async () => {
  try {
    const file = await window.api.selectDbFile();
    if (file) {
      loadRoutes();
      loadTimestampRange();
    }
  } catch (e) {
    console.error('选择数据库出错', e);
    alert('打开数据库失败：' + e.message);
  }
});

// Miss rate page: Select database
btnSelectDbMissrate.addEventListener('click', async () => {
  try {
    const file = await window.api.selectDbFile();
    if (file) {
      loadRoutesMissrate();
      loadTimestampRange();
    }
  } catch (e) {
    console.error('选择数据库出错', e);
    alert('打开数据库失败：' + e.message);
  }
});

async function loadRoutes() {
  try {
    const routes = await window.api.getRoutes();
    routeSelect.innerHTML = '<option value="">--请选择--</option>';
    routes.forEach(r => {
      const opt = document.createElement('option');
      opt.value = r.routeId;
      opt.textContent = r.routeName;
      routeSelect.appendChild(opt);
    });
  } catch (e) {
    console.error('加载路线失败', e);
    alert('加载路线失败：' + e.message);
  }
}

async function loadRoutesMissrate() {
  try {
    const routes = await window.api.getRoutes();
    routeSelectMissrate.innerHTML = '<option value="">--请选择--</option>';
    routes.forEach(r => {
      const opt = document.createElement('option');
      opt.value = r.routeId;
      opt.textContent = r.routeName;
      routeSelectMissrate.appendChild(opt);
    });
  } catch (e) {
    console.error('加载路线失败', e);
    alert('加载路线失败：' + e.message);
  }
}

async function loadTimestampRange() {
  try {
    const range = await window.api.getTimestampRange();
    timestampRange = range;
    
    if (range.minTimestamp && range.maxTimestamp) {
      // Convert timestamps to dates
      const minDate = new Date(range.minTimestamp * 1000);
      const maxDate = new Date(range.maxTimestamp * 1000);
      
      // Format dates as YYYY-MM-DD
      const minDateStr = formatDateForInput(minDate);
      const maxDateStr = formatDateForInput(maxDate);
      
      // Set min/max attributes on date inputs
      dateInput.min = minDateStr;
      dateInput.max = maxDateStr;
      
      startDateInput.min = minDateStr;
      startDateInput.max = maxDateStr;
      endDateInput.min = minDateStr;
      endDateInput.max = maxDateStr;
      
      // Set default values
      if (!dateInput.value) dateInput.value = minDateStr;
      if (!startDateInput.value) startDateInput.value = minDateStr;
      if (!endDateInput.value) endDateInput.value = maxDateStr;
    }
  } catch (e) {
    console.error('加载时间戳范围失败', e);
  }
}

function formatDateForInput(date) {
  // Convert to Beijing timezone string format
  const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Shanghai' };
  const parts = date.toLocaleDateString('zh-CN', options).split('/');
  // zh-CN format is YYYY/MM/DD
  return parts.join('-');
}

let currentRouteId = null;
let currentDate = null;

// Detail page: Preview button
btnPreview.addEventListener('click', async () => {
  const routeId = routeSelect.value;
  const date = dateInput.value;
  if (!routeId || !date) {
    alert('请选择路线和日期');
    return;
  }
  try {
    const rows = await window.api.queryRecords({ routeId, date });
    if (rows.length === 0) {
      alert('指定日期没有点检记录');
    }
    currentRecords = rows;
    currentRouteId = routeId;
    currentDate = date;
    renderTable(rows);
  } catch (e) {
    console.error('查询失败', e);
    alert('查询失败：' + e.message);
  }
});

// Detail page: Export button
btnExport.addEventListener('click', async () => {
  if (!currentRouteId || !currentDate) {
    alert('请选择路线和日期并预览后再导出');
    return;
  }
  const result = await window.api.selectExportPath();
  if (result.canceled) return;
  const success = await window.api.exportRecords({
    routeId: currentRouteId,
    date: currentDate,
    outputPath: result.filePath
  });
  if (success) {
    alert('导出成功：' + result.filePath);
  } else {
    alert('导出失败');
  }
});

// Miss rate page: Query button
btnQueryMissrate.addEventListener('click', async () => {
  const routeId = routeSelectMissrate.value;
  const startDate = startDateInput.value;
  const endDate = endDateInput.value;
  
  if (!routeId || !startDate || !endDate) {
    alert('请选择路线和起止日期');
    return;
  }
  
  if (new Date(startDate) > new Date(endDate)) {
    alert('起始日期不能晚于结束日期');
    return;
  }
  
  try {
    const stats = await window.api.queryMissRateStats({
      routeId,
      startDate,
      endDate
    });
    
    renderMissRateStats(stats);
  } catch (e) {
    console.error('查询漏检率失败', e);
    alert('查询漏检率失败：' + e.message);
  }
});

function renderMissRateStats(stats) {
  // Render left table (shift statistics)
  // Order: date, shift, value, inspected, uninspected, missRate
  missrateLeftTbody.innerHTML = '';
  stats.leftData.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.date}</td>
      <td>${row.shift}</td>
      <td>${row.value}</td>
      <td>${row.inspected}</td>
      <td>${row.uninspected}</td>
      <td>${row.missRate}</td>
    `;
    missrateLeftTbody.appendChild(tr);
  });
  
  // Render right table (value statistics)
  missrateRightTbody.innerHTML = '';
  stats.rightData.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.value}</td>
      <td>${row.inspected}</td>
      <td>${row.uninspected}</td>
      <td>${row.missRate}</td>
    `;
    missrateRightTbody.appendChild(tr);
  });
  
  refreshResizable();
}

function renderTable(rows) {
  const shiftOrder = ['夜班','白班','中班'];
  rows.sort((a,b) => {
    const s = shiftOrder.indexOf(a.shiftName) - shiftOrder.indexOf(b.shiftName);
    if (s !== 0) return s;
    return a.pointId.localeCompare(b.pointId);
  });

  tbody.innerHTML = '';
  rows.forEach(r => {
    const tr = document.createElement('tr');
    tr.style.cursor = 'pointer';
    tr.addEventListener('click', () => { showDetails(r.shiftName, r.pointId, r.slotIndex); });

    const keys = ['date','routeName','employeeName','calcValue','shiftName','pointId','pointName','freqHours','slotIndex'];
    keys.forEach(key => {
      const td = document.createElement('td');
      td.textContent = r[key] || '';
      tr.appendChild(td);
    });
    
    const statusTd = document.createElement('td');
    if (r.inspectionStatus === '已检') {
      statusTd.textContent = '已检';
    } else {
      statusTd.textContent = '未检';
      statusTd.style.color = 'red';
    }
    tr.appendChild(statusTd);
    tbody.appendChild(tr);
  });
  
  refreshResizable();
}

async function showDetails(shiftName, pointId, slotIndex) {
  if (!currentRouteId || !currentDate) return;
  console.log('request details', {shiftName, pointId, slotIndex});
  const details = await window.api.getRecordDetails({
    routeId: currentRouteId,
    date: currentDate,
    shiftName,
    pointId,
    slotIndex
  });
  console.log('details received', details);
  populateDetailModal(details);
}

function populateDetailModal(details) {
  const tbody2 = document.querySelector('#detailTable tbody');
  tbody2.innerHTML = '';
  details.forEach(d => {
    const tr = document.createElement('tr');
    ['date','time','pointId','pointName','equipmentName','itemId','itemName','freqHours','slotIndex','value','abnormal'].forEach(key => {
      const td = document.createElement('td');
      if (key === 'abnormal') {
        if (d.abnormal === 0) {
          td.textContent = '正常';
          td.style.color = '';
        } else if (d.abnormal === 1) {
          td.textContent = '异常';
          td.style.color = 'red';
        } else {
          td.textContent = d.abnormal || '';
        }
      } else {
        td.textContent = d[key] || '';
      }
      tr.appendChild(td);
    });
    tbody2.appendChild(tr);
  });
  
  refreshResizable();
  document.getElementById('detailModal').style.display = 'block';
}

// Close modal
const closeBtn = document.getElementById('closeDetail');
if (closeBtn) closeBtn.addEventListener('click', () => {
  document.getElementById('detailModal').style.display = 'none';
});

// Menu handling
const menuLinks = ['menu-detail','menu-missrate','menu-items','menu-employees'];
function setActiveMenu(id) {
  menuLinks.forEach(m => {
    const el = document.getElementById(m);
    if (el) {
      if (m === id) el.classList.add('active');
      else el.classList.remove('active');
    }
  });
}

menuLinks.forEach(id => {
  const el = document.getElementById(id);
  if (!el) return;
  el.addEventListener('click', e => {
    e.preventDefault();
    setActiveMenu(id);
    
    if (id === 'menu-detail') {
      showPage('detail');
    } else if (id === 'menu-missrate') {
      showPage('missrate');
    } else {
      alert('功能尚未实现');
    }
  });
});

// Mark default active
setActiveMenu('menu-detail');

// Column resizing helper
function makeColumnsResizable(table) {
  if (!table) return;
  const ths = table.querySelectorAll('th');
  ths.forEach(th => {
    if (th.querySelector('.resize-handle')) return;
    const handle = document.createElement('div');
    handle.className = 'resize-handle';
    th.appendChild(handle);

    let startX, startWidth;
    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      startX = e.pageX;
      startWidth = th.offsetWidth;
      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    });
    
    function onMouseMove(e) {
      const diff = e.pageX - startX;
      th.style.width = startWidth + diff + 'px';
    }
    
    function onMouseUp() {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    }
  });
}

function refreshResizable() {
  makeColumnsResizable(document.getElementById('results-table'));
  makeColumnsResizable(document.getElementById('detailTable'));
  makeColumnsResizable(document.getElementById('missrate-left-table'));
  makeColumnsResizable(document.getElementById('missrate-right-table'));
}

