/* =====================================================================
   QC Dashboard — app.js
   เก็บข้อมูลทั้งหมดใน localStorage ของเบราว์เซอร์ (ไม่มีเซิร์ฟเวอร์ภายนอก)
   โครงสร้างข้อมูลอ้างอิงตามไฟล์ Excel งาน QC Dashboard (PK / RM)
   ===================================================================== */
(function () {
  'use strict';

  var PK_SCHEMA = [
    { key: 'year',        label: 'ปี',                type: 'text',   aliases: ['ปี', 'year'] },
    { key: 'receiveNo',   label: 'เลขที่รับเข้า',     type: 'text',   aliases: ['เลขที่รับเข้า', 'receive_no', 'receiveno', 'เลขที่รับ', 'received_no'] },
    { key: 'inspectDate', label: 'วันที่ตรวจสอบ',     type: 'date',   aliases: ['วันที่ตรวจสอบ', 'inspect_date', 'inspectdate'] },
    { key: 'receiveDate', label: 'วันที่รับเข้า',      type: 'date',   aliases: ['วันที่รับเข้า', 'receive_date', 'receivedate'] },
    { key: 'retestDate',  label: 'Retest Date',       type: 'date',   aliases: ['retest date', 'retest_date', 'retestdate'] },
    { key: 'code',        label: 'รหัส',              type: 'text',   aliases: ['รหัส', 'code', 'item_code'] },
    { key: 'name',        label: 'ชื่อบรรจุภัณฑ์',    type: 'text',   aliases: ['ชื่อบรรจุภัณฑ์', 'name', 'item_name', 'ชื่อสินค้า'] },
    { key: 'supplier',    label: 'Supplier',          type: 'text',   aliases: ['supplier', 'ผู้ขาย'] },
    { key: 'lot',         label: 'Lot',               type: 'text',   aliases: ['lot', 'lotno', 'lot_no'] },
    { key: 'qty',         label: 'จำนวน(ชิ้น)',       type: 'number', aliases: ['จำนวน(ชิ้น)', 'จำนวน', 'qty', 'quantity'] },
    { key: 'status',      label: 'สถานะ',             type: 'status', statusColumns: { 'ผ่าน': 'ผ่าน', 'Quarantine': 'Quarantine', 'Rejected': 'Rejected' } },
    { key: 'approveDate', label: 'วันที่อนุมัติ',      type: 'date',   aliases: ['วันที่อนุมัติ', 'approve_date'] },
    { key: 'inspector',   label: 'ผู้ตรวจ',           type: 'text',   aliases: ['ผู้ตรวจ', 'inspector'] },
    { key: 'approver',    label: 'ผู้อนุมัติ',        type: 'text',   aliases: ['ผู้อนุมัติ', 'approver'] },
    { key: 'remark',      label: 'หมายเหตุ',          type: 'text',   aliases: ['หมายเหตุ', 'remark'] },
    { key: 'defectType',  label: 'ประเภท Defect',     type: 'text',   aliases: ['ประเภท defect (auto)', 'ประเภทdefect', 'defect', 'defect_type', 'defecttype'] }
  ];

  var RM_SCHEMA = [
    { key: 'receiveNo',   label: 'Received_no',       type: 'text',    aliases: ['received_no', 'receivedno', 'เลขที่รับเข้า'] },
    { key: 'receiveDate', label: 'วันที่รับเข้า',      type: 'date',    aliases: ['วันที่รับเข้า', 'receive_date', 'receivedate'] },
    { key: 'retestDate',  label: 'Retest_Date',       type: 'date',    aliases: ['retest_date', 'retest date', 'retestdate'] },
    { key: 'code',        label: 'รหัสวัตถุดิบ',      type: 'text',    aliases: ['รหัสวัตถุดิบ', 'code', 'rm_code'] },
    { key: 'name',        label: 'ชื่อวัตถุดิบ',      type: 'text',    aliases: ['ชื่อวัตถุดิบ', 'name', 'rm_name'] },
    { key: 'supplier',    label: 'Supplier',          type: 'text',    aliases: ['supplier', 'ผู้ขาย'] },
    { key: 'lot',         label: 'Lot',               type: 'text',    aliases: ['lot', 'lotno', 'lot_no'] },
    { key: 'deliveryNo',  label: 'เลขที่ใบส่งสินค้า',  type: 'text',    aliases: ['เลขที่ใบส่งสินค้า'] },
    { key: 'receiver',    label: 'ผู้รับสินค้า',      type: 'text',    aliases: ['ผู้รับสินค้า'] },
    { key: 'expDate',     label: 'Exp.Date',          type: 'date',    aliases: ['exp.date', 'expdate', 'exp_date'] },
    { key: 'coa',         label: 'COA',               type: 'boolean', aliases: ['coa'] },
    { key: 'status',      label: 'สถานะอนุมัติ',      type: 'status',  statusColumns: { 'Approved': 'Approved', 'Quarantine': 'Quarantine', 'Rejected': 'Rejected' } },
    { key: 'retain',      label: 'เก็บตัวอย่าง Retain', type: 'boolean', aliases: ['retain'] },
    { key: 'approveDate', label: 'วันที่อนุมัติ',      type: 'date',    aliases: ['วันที่อนุมัติ', 'approve_date'] },
    { key: 'inspector',   label: 'ผู้ตรวจ',           type: 'text',    aliases: ['ผู้ตรวจ', 'inspector'] },
    { key: 'approver',    label: 'ผู้อนุมัติ',        type: 'text',    aliases: ['ผู้อนุมัติ', 'approver'] },
    { key: 'remark',      label: 'หมายเหตุ',          type: 'text',    aliases: ['หมายเหตุ', 'remark'] },
    { key: 'remark2',     label: 'Remark',            type: 'text',    aliases: ['remark'] }
  ];

  var PK_STATUSES = ['ผ่าน', 'Quarantine', 'Rejected'];
  var RM_STATUSES = ['Approved', 'Quarantine', 'Rejected'];

  var STATUS_COLOR_HEX = { pass: '#1B998B', warn: '#E8A33D', danger: '#D7263D', retain: '#5E60CE', neutral: '#5C6784' };

  var THAI_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

  var STORAGE_KEYS = { PK: 'qc_pk_records', RM: 'qc_rm_records' };

  function loadRecords(type) {
    try {
      var raw = localStorage.getItem(STORAGE_KEYS[type]);
      var arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch (err) {
      console.error('โหลดข้อมูลไม่สำเร็จ:', err);
      return [];
    }
  }

  function saveRecords(type, arr) {
    try {
      localStorage.setItem(STORAGE_KEYS[type], JSON.stringify(arr));
      return true;
    } catch (err) {
      console.error('บันทึกข้อมูลไม่สำเร็จ:', err);
      showToast('บันทึกข้อมูลไม่สำเร็จ (พื้นที่จัดเก็บอาจเต็ม)', 'danger');
      return false;
    }
  }

  function genId() {
    return 'id_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8);
  }

  function getSchema(type) { return type === 'PK' ? PK_SCHEMA : RM_SCHEMA; }
  function getStatusList(type) { return type === 'PK' ? PK_STATUSES : RM_STATUSES; }

  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function debounce(fn, wait) {
    var t;
    return function () {
      var args = arguments, ctx = this;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(ctx, args); }, wait);
    };
  }

  function formatDateDisplay(iso) {
    if (!iso) return '-';
    var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return escapeHtml(iso);
    return m[3] + '/' + m[2] + '/' + m[1];
  }

  function todayStr() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function statusInfo(type, status) {
    var map = {
      PK: {
        'ผ่าน':       { cls: 'status-pass',   icon: 'bi-check-circle-fill',       color: 'pass' },
        'Quarantine': { cls: 'status-warn',   icon: 'bi-exclamation-triangle-fill', color: 'warn' },
        'Rejected':   { cls: 'status-danger', icon: 'bi-x-circle-fill',           color: 'danger' }
      },
      RM: {
        'Approved':   { cls: 'status-pass',   icon: 'bi-check-circle-fill',       color: 'pass' },
        'Quarantine': { cls: 'status-warn',   icon: 'bi-exclamation-triangle-fill', color: 'warn' },
        'Rejected':   { cls: 'status-danger', icon: 'bi-x-circle-fill',           color: 'danger' }
      }
    };
    return (map[type] && map[type][status]) || { cls: 'status-warn', icon: 'bi-question-circle', color: 'neutral' };
  }

  function statusBadgeHTML(type, status) {
    var info = statusInfo(type, status);
    return '<span class="status-badge ' + info.cls + '"><i class="bi ' + info.icon + '"></i>' + escapeHtml(status || '-') + '</span>';
  }

  function showToast(message, variant) {
    variant = variant || 'success';
    var iconMap = { success: 'bi-check-circle-fill', danger: 'bi-x-circle-fill', warning: 'bi-exclamation-triangle-fill' };
    var bgMap = { success: 'text-bg-success', danger: 'text-bg-danger', warning: 'text-bg-warning' };
    var el = document.createElement('div');
    el.className = 'toast align-items-center ' + (bgMap[variant] || 'text-bg-secondary');
    el.setAttribute('role', 'alert');
    el.innerHTML =
      '<div class="d-flex">' +
      '<div class="toast-body"><i class="bi ' + (iconMap[variant] || 'bi-info-circle-fill') + ' me-2"></i>' + escapeHtml(message) + '</div>' +
      '<button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>' +
      '</div>';
    document.getElementById('toastContainer').appendChild(el);
    var toast = new bootstrap.Toast(el, { delay: 3200 });
    toast.show();
    el.addEventListener('hidden.bs.toast', function () { el.remove(); });
  }

  function askConfirm(message, onConfirm) {
    document.getElementById('confirmModalBody').textContent = message;
    var modalEl = document.getElementById('confirmModal');
    var modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    var okBtn = document.getElementById('confirmModalOkBtn');
    okBtn.onclick = function () {
      modal.hide();
      onConfirm();
    };
    modal.show();
  }

  var PAGES = ['dashboard', 'entry', 'table', 'daily', 'analysis', 'io'];
  var TITLES = {
    dashboard: 'ภาพรวม (Dashboard)',
    entry: 'บันทึกข้อมูล',
    table: 'ตารางข้อมูล / ค้นหา',
    daily: 'สรุปรายวัน (Daily Summary)',
    analysis: 'วิเคราะห์คุณภาพ (Quality Analysis)',
    io: 'นำเข้า / ส่งออก Excel'
  };
  var activePage = 'dashboard';

  function gotoPage(page) {
    if (PAGES.indexOf(page) === -1) page = 'dashboard';
    activePage = page;
    PAGES.forEach(function (p) {
      document.getElementById('page-' + p).classList.toggle('d-none', p !== page);
    });
    document.querySelectorAll('.sidebar-link').forEach(function (a) {
      a.classList.toggle('active', a.dataset.page === page);
    });
    document.getElementById('pageTitle').textContent = TITLES[page];
    history.replaceState(null, '', '#' + page);

    if (page === 'dashboard') renderDashboard();
    if (page === 'table') { renderTableHead(); populateStatusFilterOptions(); renderTableBody(); }
    if (page === 'daily') renderDailySummary();
    if (page === 'analysis') renderAnalysis();

    var sidebarEl = document.getElementById('sidebar');
    var oc = bootstrap.Offcanvas.getInstance(sidebarEl);
    if (oc) oc.hide();
  }

  var chartInstances = { PK: null, RM: null };

  function computeKPI(type, data) {
    var statuses = getStatusList(type);
    var counts = {};
    statuses.forEach(function (s) { counts[s] = 0; });
    data.forEach(function (r) { if (counts.hasOwnProperty(r.status)) counts[r.status]++; });
    var total = data.length;
    var knownTotal = statuses.reduce(function (sum, s) { return sum + counts[s]; }, 0);
    var passKey = type === 'PK' ? 'ผ่าน' : 'Approved';
    var passRate = knownTotal ? ((counts[passKey] / knownTotal) * 100).toFixed(1) : '0.0';
    return { total: total, knownTotal: knownTotal, counts: counts, passRate: passRate };
  }

  function renderStatusChart(canvasId, emptyId, type, data) {
    var statuses = getStatusList(type);
    var kpi = computeKPI(type, data);
    var canvas = document.getElementById(canvasId);
    var emptyEl = document.getElementById(emptyId);

    if (kpi.total === 0) {
      canvas.classList.add('d-none');
      emptyEl.classList.remove('d-none');
      if (chartInstances[type]) { chartInstances[type].destroy(); chartInstances[type] = null; }
      return;
    }
    canvas.classList.remove('d-none');
    emptyEl.classList.add('d-none');

    var values = statuses.map(function (s) { return kpi.counts[s]; });
    var colors = statuses.map(function (s) { return STATUS_COLOR_HEX[statusInfo(type, s).color]; });

    if (chartInstances[type]) chartInstances[type].destroy();
    var inkColor = getComputedStyle(document.documentElement).getPropertyValue('--color-ink').trim() || '#1C2541';
    chartInstances[type] = new Chart(canvas.getContext('2d'), {
      type: 'doughnut',
      data: { labels: statuses, datasets: [{ data: values, backgroundColor: colors, borderWidth: 0 }] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '62%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 10, color: inkColor, font: { family: 'Sarabun', size: 11 } } }
        }
      }
    });
  }

  function renderRecentList(containerId, type, data) {
    var container = document.getElementById(containerId);
    var recent = data.slice(-5).reverse();
    if (recent.length === 0) {
      container.innerHTML = '<div class="text-muted small">ยังไม่มีข้อมูล</div>';
      return;
    }
    container.innerHTML = recent.map(function (r) {
      return '<div class="recent-item">' +
        '<div><div class="ri-name">' + escapeHtml(r.name || '-') + '</div>' +
        '<div class="ri-sub">Lot ' + escapeHtml(r.lot || '-') + ' • ' + escapeHtml(r.supplier || '-') + '</div></div>' +
        statusBadgeHTML(type, r.status) +
        '</div>';
    }).join('');
  }

  function renderDashboard() {
    var pk = loadRecords('PK');
    var rm = loadRecords('RM');

    document.getElementById('dashboardEmpty').classList.toggle('d-none', !(pk.length === 0 && rm.length === 0));
    document.getElementById('dashboardContent').classList.toggle('d-none', (pk.length === 0 && rm.length === 0));

    var pkKpi = computeKPI('PK', pk);
    document.getElementById('pkTotal').textContent = pkKpi.total;
    document.getElementById('pkPass').textContent = pkKpi.counts['ผ่าน'];
    document.getElementById('pkRejected').textContent = pkKpi.counts['Rejected'];
    document.getElementById('pkPassRate').textContent = pkKpi.passRate;

    var rmKpi = computeKPI('RM', rm);
    document.getElementById('rmTotal').textContent = rmKpi.total;
    document.getElementById('rmApproved').textContent = rmKpi.counts['Approved'];
    document.getElementById('rmRejected').textContent = rmKpi.counts['Rejected'];
    document.getElementById('rmPassRate').textContent = rmKpi.passRate;

    renderStatusChart('pkChart', 'pkChartEmpty', 'PK', pk);
    renderStatusChart('rmChart', 'rmChartEmpty', 'RM', rm);
    renderRecentList('pkRecentList', 'PK', pk);
    renderRecentList('rmRecentList', 'RM', rm);
  }

  /* ---------------- หน้าสรุปการรับเข้ารายวัน (Daily Summary) ---------------- */
  var dailyState = { year: new Date().getFullYear(), month: new Date().getMonth() + 1 };

  function getDaysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
  }

  function parseDateParts(iso) {
    var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return null;
    return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
  }

  function computeDailyCounts(type, year, month) {
    var records = loadRecords(type);
    var passKey = type === 'PK' ? 'ผ่าน' : 'Approved';
    var daysInMonth = getDaysInMonth(year, month);
    var rows = [];
    for (var d = 1; d <= daysInMonth; d++) rows.push({ day: d, total: 0, pass: 0, rejected: 0 });

    records.forEach(function (r) {
      var parts = parseDateParts(r.receiveDate);
      if (!parts || parts.y !== year || parts.m !== month) return;
      var row = rows[parts.d - 1];
      if (!row) return;
      row.total++;
      if (r.status === passKey) row.pass++;
      if (r.status === 'Rejected') row.rejected++;
    });
    return rows;
  }

  function initDailySummaryControls() {
    var monthSel = document.getElementById('dailyMonthSelect');
    var yearSel = document.getElementById('dailyYearSelect');
    if (!monthSel || !yearSel) return;

    monthSel.innerHTML = THAI_MONTHS.map(function (name, idx) {
      return '<option value="' + (idx + 1) + '">' + escapeHtml(name) + '</option>';
    }).join('');

    var currentYear = new Date().getFullYear();
    var years = [];
    for (var y = currentYear - 3; y <= currentYear + 1; y++) years.push(y);
    yearSel.innerHTML = years.map(function (y) {
      return '<option value="' + y + '">พ.ศ. ' + (y + 543) + '</option>';
    }).join('');

    monthSel.value = String(dailyState.month);
    yearSel.value = String(dailyState.year);

    monthSel.addEventListener('change', function () {
      dailyState.month = Number(this.value);
      renderDailySummary();
    });
    yearSel.addEventListener('change', function () {
      dailyState.year = Number(this.value);
      renderDailySummary();
    });
  }

  function renderDailySummary() {
    var year = dailyState.year, month = dailyState.month;
    var pkRows = computeDailyCounts('PK', year, month);
    var rmRows = computeDailyCounts('RM', year, month);
    var daysInMonth = pkRows.length;

    var totals = { pkIn: 0, pkPass: 0, pkRej: 0, rmIn: 0, rmPass: 0, rmRej: 0 };
    var bodyHtml = '';

    for (var i = 0; i < daysInMonth; i++) {
      var day = i + 1;
      var pk = pkRows[i], rm = rmRows[i];
      var dow = new Date(year, month - 1, day).getDay();
      var isWeekend = (dow === 0 || dow === 6);
      var dateLabel = String(day).padStart(2, '0') + '/' + String(month).padStart(2, '0') + '/' + year;

      totals.pkIn += pk.total; totals.pkPass += pk.pass; totals.pkRej += pk.rejected;
      totals.rmIn += rm.total; totals.rmPass += rm.pass; totals.rmRej += rm.rejected;

      bodyHtml += '<tr' + (isWeekend ? ' class="weekend-row"' : '') + '>' +
        '<td class="mono">' + day + '</td>' +
        '<td class="mono">' + dateLabel + '</td>' +
        '<td class="mono">' + pk.total + '</td>' +
        '<td class="mono">' + pk.pass + '</td>' +
        '<td class="mono">' + pk.rejected + '</td>' +
        '<td class="mono">' + rm.total + '</td>' +
        '<td class="mono">' + rm.pass + '</td>' +
        '<td class="mono">' + rm.rejected + '</td>' +
        '</tr>';
    }

    var bodyEl = document.getElementById('dailySummaryBody');
    if (bodyEl) bodyEl.innerHTML = bodyHtml;

    var setText = function (id, val) {
      var el = document.getElementById(id);
      if (el) el.textContent = val;
    };
    setText('dailyTotalPkIn', totals.pkIn);
    setText('dailyTotalPkPass', totals.pkPass);
    setText('dailyTotalPkRej', totals.pkRej);
    setText('dailyTotalRmIn', totals.rmIn);
    setText('dailyTotalRmPass', totals.rmPass);
    setText('dailyTotalRmRej', totals.rmRej);
  }

  /* ---------------- หน้าวิเคราะห์คุณภาพ: สรุป Defect / ประเมิน Supplier ---------------- */
  var analysisState = { tab: 'defect', type: 'PK', supplierSort: { field: 'passRate', dir: 'asc' } };

  function getPassKey(type) { return type === 'PK' ? 'ผ่าน' : 'Approved'; }

  // ดึงข้อความ Defect: PK ใช้ "ประเภท Defect" ก่อน ถ้าไม่มีใช้ "หมายเหตุ"; RM ใช้ "หมายเหตุ" ก่อน ถ้าไม่มีใช้ "Remark"
  function getDefectText(type, r) {
    if (type === 'PK') {
      var dt = (r.defectType || '').toString().trim();
      if (dt) return dt;
      return (r.remark || '').toString().trim();
    }
    var rm1 = (r.remark || '').toString().trim();
    if (rm1) return rm1;
    return (r.remark2 || '').toString().trim();
  }

  function computeDefectSummary(type) {
    var records = loadRecords(type);
    var passKey = getPassKey(type);

    var groups = {};
    var totalTagged = 0;
    var statusBreakdown = {};

    // นับทุกรายการที่มีข้อความ Defect/หมายเหตุ ไม่ว่าสถานะสุดท้ายจะเป็นอะไร
    // (ในข้อมูลจริงมีการบันทึกตำหนิ/หมายเหตุไว้แม้รายการนั้นสถานะสุดท้ายจะ "ผ่าน" ก็ตาม)
    records.forEach(function (r) {
      var text = getDefectText(type, r);
      if (!text) return;
      totalTagged++;

      var st = r.status || 'ไม่ระบุสถานะ';
      statusBreakdown[st] = (statusBreakdown[st] || 0) + 1;

      if (!groups[text]) groups[text] = { text: text, count: 0, suppliers: {} };
      groups[text].count++;
      var sup = (r.supplier || '').toString().trim() || 'ไม่ระบุ';
      groups[text].suppliers[sup] = true;
    });

    var list = Object.keys(groups).map(function (k) {
      var g = groups[k];
      return { text: g.text, count: g.count, supplierCount: Object.keys(g.suppliers).length };
    });
    list.sort(function (a, b) { return b.count - a.count; });

    return { totalTagged: totalTagged, uniqueCount: list.length, top: list[0] || null, list: list, statusBreakdown: statusBreakdown, passKey: passKey };
  }

  var defectChartInstance = null;

  function renderDefectPanel() {
    var summary = computeDefectSummary(analysisState.type);

    document.getElementById('defectTotalProblem').textContent = summary.totalTagged;
    document.getElementById('defectUniqueCount').textContent = summary.uniqueCount;
    document.getElementById('defectTopName').textContent = summary.top ? summary.top.text : '-';
    document.getElementById('defectTopCount').textContent = (summary.top ? summary.top.count : 0) + ' ครั้ง';

    var breakdownEl = document.getElementById('defectStatusBreakdown');
    if (breakdownEl) {
      var parts = Object.keys(summary.statusBreakdown).map(function (st) {
        var displayLabel = (st === summary.passKey) ? (st + ' (ผ่านแต่มีบันทึกไว้)') : st;
        return displayLabel + ' ' + summary.statusBreakdown[st] + ' รายการ';
      });
      breakdownEl.textContent = parts.length ? ('แยกตามสถานะ: ' + parts.join(' · ')) : '';
    }

    var tbody = document.getElementById('defectTableBody');
    var tableEl = document.getElementById('defectTable');
    var tableEmptyEl = document.getElementById('defectTableEmpty');
    if (summary.list.length === 0) {
      tbody.innerHTML = '';
      tableEl.classList.add('d-none');
      tableEmptyEl.classList.remove('d-none');
    } else {
      tableEl.classList.remove('d-none');
      tableEmptyEl.classList.add('d-none');
      tbody.innerHTML = summary.list.map(function (g) {
        var pct = summary.totalTagged ? ((g.count / summary.totalTagged) * 100).toFixed(1) : '0.0';
        return '<tr><td>' + escapeHtml(g.text) + '</td><td class="mono">' + g.count + '</td><td class="mono">' + pct + '%</td><td class="mono">' + g.supplierCount + '</td></tr>';
      }).join('');
    }

    var canvas = document.getElementById('defectChart');
    var chartEmptyEl = document.getElementById('defectChartEmpty');
    var top8 = summary.list.slice(0, 8);

    if (top8.length === 0) {
      canvas.classList.add('d-none');
      chartEmptyEl.classList.remove('d-none');
      if (defectChartInstance) { defectChartInstance.destroy(); defectChartInstance = null; }
    } else {
      canvas.classList.remove('d-none');
      chartEmptyEl.classList.add('d-none');
      if (defectChartInstance) defectChartInstance.destroy();
      var inkColor = getComputedStyle(document.documentElement).getPropertyValue('--color-ink').trim() || '#1C2541';
      defectChartInstance = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
          labels: top8.map(function (g) { return g.text.length > 22 ? g.text.slice(0, 22) + '…' : g.text; }),
          datasets: [{ data: top8.map(function (g) { return g.count; }), backgroundColor: STATUS_COLOR_HEX.danger, borderRadius: 4 }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { ticks: { color: inkColor, precision: 0 }, grid: { color: 'rgba(128,128,128,.12)' } },
            y: { ticks: { color: inkColor, font: { family: 'Sarabun', size: 11 } }, grid: { display: false } }
          }
        }
      });
    }
  }

  function gradeFromRate(rate) {
    if (rate >= 98) return { label: 'A', cls: 'status-pass' };
    if (rate >= 95) return { label: 'B', cls: 'status-pass' };
    if (rate >= 90) return { label: 'C', cls: 'status-warn' };
    return { label: 'D', cls: 'status-danger' };
  }

  function computeSupplierEvaluation(type) {
    var records = loadRecords(type);
    var passKey = getPassKey(type);
    var groups = {};
    records.forEach(function (r) {
      var sup = (r.supplier || '').toString().trim() || 'ไม่ระบุ Supplier';
      if (!groups[sup]) groups[sup] = { supplier: sup, total: 0, pass: 0, quarantine: 0, rejected: 0 };
      var g = groups[sup];
      g.total++;
      if (r.status === passKey) g.pass++;
      else if (r.status === 'Quarantine') g.quarantine++;
      else if (r.status === 'Rejected') g.rejected++;
    });
    return Object.keys(groups).map(function (k) {
      var g = groups[k];
      g.passRate = g.total ? (g.pass / g.total) * 100 : 0;
      g.grade = gradeFromRate(g.passRate).label;
      return g;
    });
  }

  function sortSupplierList(list) {
    var field = analysisState.supplierSort.field;
    var dir = analysisState.supplierSort.dir === 'asc' ? 1 : -1;
    list.sort(function (a, b) {
      var av = a[field], bv = b[field];
      if (typeof av === 'string') { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      if (av < bv) return -1 * dir;
      if (av > bv) return 1 * dir;
      return 0;
    });
    return list;
  }

  function renderSupplierPanel() {
    var list = sortSupplierList(computeSupplierEvaluation(analysisState.type));

    var tbody = document.getElementById('supplierTableBody');
    var tableEl = document.getElementById('supplierTable');
    var emptyEl = document.getElementById('supplierTableEmpty');

    if (list.length === 0) {
      tbody.innerHTML = '';
      tableEl.classList.add('d-none');
      emptyEl.classList.remove('d-none');
    } else {
      tableEl.classList.remove('d-none');
      emptyEl.classList.add('d-none');
      tbody.innerHTML = list.map(function (g) {
        var grade = gradeFromRate(g.passRate);
        return '<tr>' +
          '<td>' + escapeHtml(g.supplier) + '</td>' +
          '<td class="mono">' + g.total + '</td>' +
          '<td class="mono">' + g.pass + '</td>' +
          '<td class="mono">' + g.quarantine + '</td>' +
          '<td class="mono">' + g.rejected + '</td>' +
          '<td class="mono">' + g.passRate.toFixed(1) + '%</td>' +
          '<td><span class="status-badge grade-badge ' + grade.cls + '">' + grade.label + '</span></td>' +
          '</tr>';
      }).join('');
    }

    document.querySelectorAll('#supplierTable th.sortable').forEach(function (th) {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.dataset.sort === analysisState.supplierSort.field) {
        th.classList.add(analysisState.supplierSort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      }
    });
  }

  function renderAnalysis() {
    if (analysisState.tab === 'defect') renderDefectPanel();
    else renderSupplierPanel();
  }

  function handleEntrySubmit(e, type) {
    e.preventDefault();
    var form = e.target;
    if (!form.checkValidity()) {
      form.classList.add('was-validated');
      return;
    }
    var fd = new FormData(form);
    var schema = getSchema(type);
    var record = {};
    schema.forEach(function (f) {
      var val = fd.get(f.key);
      if (f.type === 'number') val = (val === '' || val === null) ? '' : Number(val);
      if (f.type === 'boolean') val = (val === 'on');
      record[f.key] = (val === null) ? '' : val;
    });
    record.id = genId();
    record.createdAt = Date.now();

    var arr = loadRecords(type);
    arr.push(record);
    saveRecords(type, arr);

    form.reset();
    form.classList.remove('was-validated');
    showToast('บันทึกข้อมูล ' + type + ' เรียบร้อยแล้ว', 'success');
    refreshAll();
  }

  var PAGE_SIZE = 50;
  var tableState = { currentTab: 'PK', searchQuery: '', statusFilter: '', currentPage: 1 };

  function renderTableHead() {
    var schema = getSchema(tableState.currentTab);
    var row = document.getElementById('tableHeadRow');
    row.innerHTML = schema.map(function (f) { return '<th>' + escapeHtml(f.label) + '</th>'; }).join('') + '<th>การจัดการ</th>';
  }

  function populateStatusFilterOptions() {
    var sel = document.getElementById('statusFilter');
    var statuses = getStatusList(tableState.currentTab);
    sel.innerHTML = '<option value="">ทุกสถานะ</option>' +
      statuses.map(function (s) { return '<option value="' + escapeHtml(s) + '">' + escapeHtml(s) + '</option>'; }).join('');
    sel.value = '';
    tableState.statusFilter = '';
  }

  function getFilteredData() {
    var data = loadRecords(tableState.currentTab);
    var q = tableState.searchQuery.trim().toLowerCase();
    return data.filter(function (r) {
      var matchesQuery = !q || [r.name, r.lot, r.supplier, r.code].some(function (v) {
        return String(v || '').toLowerCase().indexOf(q) !== -1;
      });
      var matchesStatus = !tableState.statusFilter || r.status === tableState.statusFilter;
      return matchesQuery && matchesStatus;
    });
  }

  function renderPagination(totalItems) {
    var bar = document.getElementById('paginationBar');
    var totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
    if (tableState.currentPage > totalPages) tableState.currentPage = totalPages;
    if (tableState.currentPage < 1) tableState.currentPage = 1;

    bar.classList.toggle('d-none', totalItems === 0);
    document.getElementById('pageInfo').textContent = 'หน้า ' + tableState.currentPage + ' / ' + totalPages;
    document.getElementById('prevPageBtn').disabled = (tableState.currentPage <= 1);
    document.getElementById('nextPageBtn').disabled = (tableState.currentPage >= totalPages);
  }

  function renderTableBody() {
    var type = tableState.currentTab;
    var schema = getSchema(type);
    var totalCount = loadRecords(type).length;
    var filtered = getFilteredData();

    document.getElementById('resultCount').textContent = filtered.length + ' / ' + totalCount + ' รายการ';

    var tbody = document.getElementById('tableBody');
    var tableEl = document.getElementById('dataTable');
    var emptyEl = document.getElementById('tableEmpty');

    if (filtered.length === 0) {
      tbody.innerHTML = '';
      tableEl.classList.add('d-none');
      emptyEl.classList.remove('d-none');
      renderPagination(0);
      return;
    }
    tableEl.classList.remove('d-none');
    emptyEl.classList.add('d-none');

    var reversed = filtered.slice().reverse();
    renderPagination(reversed.length);
    var startIdx = (tableState.currentPage - 1) * PAGE_SIZE;
    var pageItems = reversed.slice(startIdx, startIdx + PAGE_SIZE);

    tbody.innerHTML = pageItems.map(function (r) {
      var cells = schema.map(function (f) {
        var val = r[f.key];
        if (f.type === 'status') return '<td>' + statusBadgeHTML(type, val) + '</td>';
        if (f.type === 'date') return '<td>' + formatDateDisplay(val) + '</td>';
        if (f.type === 'boolean') return '<td>' + (val ? '<span class="status-badge status-pass"><i class="bi bi-check-circle-fill"></i>มี</span>' : '<span class="text-muted">-</span>') + '</td>';
        if (f.type === 'number') return '<td class="mono">' + ((val === '' || val === undefined || val === null) ? '-' : Number(val).toLocaleString('th-TH')) + '</td>';
        return '<td>' + (val ? escapeHtml(String(val)) : '-') + '</td>';
      }).join('');
      return '<tr>' + cells +
        '<td class="row-actions">' +
        '<button class="btn btn-sm btn-outline-secondary btn-edit" data-id="' + r.id + '" title="แก้ไข"><i class="bi bi-pencil"></i></button> ' +
        '<button class="btn btn-sm btn-outline-danger btn-delete" data-id="' + r.id + '" title="ลบ"><i class="bi bi-trash3"></i></button>' +
        '</td></tr>';
    }).join('');
  }

  function buildEditFormHTML(type, record) {
    var schema = getSchema(type);
    var html = '<form id="editForm" novalidate><div class="row g-3">';
    schema.forEach(function (f) {
      var val = record[f.key];
      val = (val === undefined || val === null) ? '' : val;
      var inputHtml;
      if (f.type === 'status') {
        var statuses = getStatusList(type);
        var opts = statuses.map(function (s) {
          return '<option value="' + escapeHtml(s) + '"' + (s === val ? ' selected' : '') + '>' + escapeHtml(s) + '</option>';
        }).join('');
        inputHtml = '<select class="form-select" name="' + f.key + '" required>' +
          '<option value=""' + (!val ? ' selected' : '') + ' disabled>เลือกสถานะ</option>' + opts + '</select>';
      } else if (f.type === 'date') {
        inputHtml = '<input type="date" class="form-control" name="' + f.key + '" value="' + escapeHtml(val) + '">';
      } else if (f.type === 'boolean') {
        inputHtml = '<div class="form-check mt-2"><input class="form-check-input" type="checkbox" name="' + f.key + '"' + (val ? ' checked' : '') + '></div>';
      } else if (f.type === 'number') {
        inputHtml = '<input type="number" min="0" class="form-control" name="' + f.key + '" value="' + escapeHtml(val) + '">';
      } else {
        inputHtml = '<input type="text" class="form-control" name="' + f.key + '" value="' + escapeHtml(val) + '">';
      }
      html += '<div class="col-md-6"><label class="form-label">' + escapeHtml(f.label) + '</label>' + inputHtml + '</div>';
    });
    html += '</div></form>';
    return html;
  }

  function openEditModal(type, id) {
    var arr = loadRecords(type);
    var record = arr.find(function (r) { return r.id === id; });
    if (!record) return;

    document.getElementById('editModalTitle').textContent = 'แก้ไขข้อมูล ' + type + (record.name ? ' — ' + record.name : '');
    document.getElementById('editModalBody').innerHTML = buildEditFormHTML(type, record);

    var modalEl = document.getElementById('editModal');
    var modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    modal.show();

    document.getElementById('saveEditBtn').onclick = function () {
      var form = document.getElementById('editForm');
      if (!form.checkValidity()) { form.classList.add('was-validated'); return; }
      var fd = new FormData(form);
      var schema = getSchema(type);
      schema.forEach(function (f) {
        var val = fd.get(f.key);
        if (f.type === 'number') val = (val === '' || val === null) ? '' : Number(val);
        if (f.type === 'boolean') val = (val === 'on');
        record[f.key] = (val === null) ? '' : val;
      });
      saveRecords(type, arr);
      modal.hide();
      showToast('บันทึกการแก้ไขเรียบร้อยแล้ว', 'success');
      refreshAll();
    };
  }

  function deleteRecord(type, id) {
    askConfirm('ต้องการลบรายการนี้ใช่หรือไม่? การลบไม่สามารถย้อนกลับได้', function () {
      var arr = loadRecords(type).filter(function (r) { return r.id !== id; });
      saveRecords(type, arr);
      showToast('ลบรายการเรียบร้อยแล้ว', 'danger');
      refreshAll();
    });
  }

  var pendingImport = { PK: [], RM: [] };
  var pendingImportFound = { PK: false, RM: false };

  function normalizeHeader(h) {
    return String(h).trim().toLowerCase().replace(/[\s_().]/g, '');
  }

  // แปลง Date object เป็น YYYY-MM-DD โดยอ่านค่าแบบ UTC เท่านั้น
  // (เดิมใช้ getTimezoneOffset() แล้วบวกชดเชยซ้ำ ทำให้วันที่เลื่อนถอยหลัง 1 วัน
  //  สำหรับทุกเครื่อง/เบราว์เซอร์ที่ตั้งโซนเวลาเป็น UTC+ เช่นประเทศไทย)
  function dateToInputValue(d) {
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }

  // แปลงเลข serial date ของ Excel (เช่น 46195) เป็น YYYY-MM-DD
  // คำนวณแบบ UTC ตรง ๆ ไม่พึ่งพาโซนเวลาของเบราว์เซอร์ผู้ใช้ เพื่อไม่ให้วันที่คลาดเคลื่อน
  function excelSerialToISO(serial) {
    var utcMs = Date.UTC(1899, 11, 30) + Math.round(serial * 86400000);
    var d = new Date(utcMs);
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }

  function normalizeDateValue(val) {
    if (val === '' || val === null || val === undefined) return '';
    if (val instanceof Date && !isNaN(val)) return dateToInputValue(val);
    if (typeof val === 'number' && isFinite(val)) return excelSerialToISO(val);
    var s = String(val).trim();
    if (!s) return '';
    if (/^(null|undefined|nan|n\/a|none|-)$/i.test(s)) return '';
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    return s;
  }

  function deriveStatus(normalizedRow, statusColumns) {
    for (var label in statusColumns) {
      var nk = normalizeHeader(statusColumns[label]);
      var val = normalizedRow[nk];
      if (val !== undefined && String(val).trim() !== '') return label;
    }
    return '';
  }

  function mapRowToRecord(type, rawRow) {
    var schema = getSchema(type);
    var normalizedRow = {};
    Object.keys(rawRow).forEach(function (k) { normalizedRow[normalizeHeader(k)] = rawRow[k]; });

    var record = {};
    schema.forEach(function (f) {
      if (f.type === 'status') {
        record[f.key] = deriveStatus(normalizedRow, f.statusColumns);
        return;
      }
      var val = '';
      for (var i = 0; i < f.aliases.length; i++) {
        var nk = normalizeHeader(f.aliases[i]);
        if (normalizedRow.hasOwnProperty(nk) && normalizedRow[nk] !== '' && normalizedRow[nk] !== undefined) {
          val = normalizedRow[nk];
          break;
        }
      }
      if (f.type === 'date') val = normalizeDateValue(val);
      else if (f.type === 'number') val = (val === '' || val === undefined) ? '' : (Number(val) || 0);
      else if (f.type === 'boolean') val = (String(val).trim() !== '');
      else val = (typeof val === 'string') ? val.trim() : (val === undefined ? '' : val);
      record[f.key] = val;
    });
    record.id = genId();
    record.createdAt = Date.now();
    return record;
  }

  function detectHeaderRowIndex(sheet, schema) {
    var rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    var keywords = [];
    schema.forEach(function (f) {
      if (f.type === 'status' && f.statusColumns) {
        Object.keys(f.statusColumns).forEach(function (label) { keywords.push(normalizeHeader(f.statusColumns[label])); });
      }
      (f.aliases || []).forEach(function (a) { keywords.push(normalizeHeader(a)); });
    });
    var bestIdx = 0, bestScore = -1;
    var scanLimit = Math.min(rawRows.length, 6);
    for (var i = 0; i < scanLimit; i++) {
      var rowCells = rawRows[i] || [];
      var score = 0;
      rowCells.forEach(function (cell) {
        var n = normalizeHeader(cell);
        if (n && keywords.indexOf(n) !== -1) score++;
      });
      if (score > bestScore) { bestScore = score; bestIdx = i; }
    }
    return bestIdx;
  }

  function findSheetName(workbook, keywords) {
    var names = workbook.SheetNames;
    for (var i = 0; i < names.length; i++) {
      var n = names[i].toLowerCase();
      if (keywords.some(function (kw) { return n.indexOf(kw) !== -1; })) return names[i];
    }
    return null;
  }

  function handlePreviewImport() {
    var fileInput = document.getElementById('importFile');
    var file = fileInput.files[0];
    var logEl = document.getElementById('importLog');
    if (!file) return;

    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        var pkSheetName = findSheetName(workbook, ['pk', 'บรรจุภัณฑ์']);
        var rmSheetName = findSheetName(workbook, ['rm', 'วัตถุดิบ']);

        pendingImport.PK = [];
        pendingImport.RM = [];
        pendingImportFound.PK = !!pkSheetName;
        pendingImportFound.RM = !!rmSheetName;
        var logLines = [];

        if (pkSheetName) {
          var pkHeaderIdx = detectHeaderRowIndex(workbook.Sheets[pkSheetName], PK_SCHEMA);
          var pkRows = XLSX.utils.sheet_to_json(workbook.Sheets[pkSheetName], { defval: '', range: pkHeaderIdx });
          pendingImport.PK = pkRows.map(function (row) { return mapRowToRecord('PK', row); });
          logLines.push('พบชีต "' + pkSheetName + '" → ' + pendingImport.PK.length + ' แถว (PK)');
          var pkBlankStatus = pendingImport.PK.filter(function (r) { return !r.status; });
          if (pkBlankStatus.length > 0) {
            logLines.push('⚠ พบ ' + pkBlankStatus.length + ' แถว (PK) ที่ไม่ได้ติ๊กสถานะใดเลย (ผ่าน/Quarantine/Rejected ว่างหมด) — เลขที่รับเข้า: ' +
              pkBlankStatus.slice(0, 10).map(function (r) { return r.receiveNo || '(ไม่มีเลขที่)'; }).join(', ') +
              (pkBlankStatus.length > 10 ? ' และอื่นๆ' : '') + ' — แถวเหล่านี้จะถูกนำเข้าแต่จะไม่ขึ้นในกราฟ/ผ่าน-Rejected จนกว่าจะแก้ไขสถานะ');
          }
        } else {
          logLines.push('ไม่พบชีตที่ตรงกับข้อมูล PK ในไฟล์นี้');
        }

        if (rmSheetName) {
          var rmHeaderIdx = detectHeaderRowIndex(workbook.Sheets[rmSheetName], RM_SCHEMA);
          var rmRows = XLSX.utils.sheet_to_json(workbook.Sheets[rmSheetName], { defval: '', range: rmHeaderIdx });
          pendingImport.RM = rmRows.map(function (row) { return mapRowToRecord('RM', row); });
          logLines.push('พบชีต "' + rmSheetName + '" → ' + pendingImport.RM.length + ' แถว (RM)');
          var rmBlankStatus = pendingImport.RM.filter(function (r) { return !r.status; });
          if (rmBlankStatus.length > 0) {
            logLines.push('⚠ พบ ' + rmBlankStatus.length + ' แถว (RM) ที่ไม่ได้ติ๊กสถานะใดเลย (Approved/Quarantine/Rejected ว่างหมด) — เลขที่รับเข้า: ' +
              rmBlankStatus.slice(0, 10).map(function (r) { return r.receiveNo || '(ไม่มีเลขที่)'; }).join(', ') +
              (rmBlankStatus.length > 10 ? ' และอื่นๆ' : '') + ' — แถวเหล่านี้จะถูกนำเข้าแต่จะไม่ขึ้นในกราฟ/Approved-Rejected จนกว่าจะแก้ไขสถานะ');
          }
        } else {
          logLines.push('ไม่พบชีตที่ตรงกับข้อมูล RM ในไฟล์นี้');
        }

        var summaryEl = document.getElementById('importSummary');
        summaryEl.innerHTML = logLines.map(function (l) { return '<div>' + escapeHtml(l) + '</div>'; }).join('');
        document.getElementById('importPreview').classList.remove('d-none');

        if (pendingImport.PK.length === 0 && pendingImport.RM.length === 0) {
          showToast('ไม่พบข้อมูลที่นำเข้าได้ในไฟล์นี้ ตรวจสอบชื่อชีตและหัวคอลัมน์', 'warning');
        }
      } catch (err) {
        console.error(err);
        showToast('อ่านไฟล์ไม่สำเร็จ ตรวจสอบว่าเป็นไฟล์ .xlsx ที่ถูกต้อง', 'danger');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleConfirmImport() {
    var mode = document.querySelector('input[name="importMode"]:checked').value;
    var addedPK = pendingImport.PK.length, addedRM = pendingImport.RM.length;

    if (pendingImportFound.PK) {
      var pkArr = mode === 'replace' ? pendingImport.PK : loadRecords('PK').concat(pendingImport.PK);
      saveRecords('PK', pkArr);
    }

    if (pendingImportFound.RM) {
      var rmArr = mode === 'replace' ? pendingImport.RM : loadRecords('RM').concat(pendingImport.RM);
      saveRecords('RM', rmArr);
    }

    var logEl = document.getElementById('importLog');
    var stamp = new Date().toLocaleString('th-TH');
    var entry = document.createElement('div');
    entry.textContent = '[' + stamp + '] นำเข้าสำเร็จ — PK ' + addedPK + ' แถว, RM ' + addedRM + ' แถว (โหมด: ' + (mode === 'replace' ? 'แทนที่ข้อมูลเดิม' : 'เพิ่มต่อจากข้อมูลเดิม') + ')';
    logEl.prepend(entry);

    document.getElementById('importPreview').classList.add('d-none');
    document.getElementById('importFile').value = '';
    document.getElementById('previewImportBtn').disabled = true;
    pendingImport = { PK: [], RM: [] };
    pendingImportFound = { PK: false, RM: false };

    showToast('นำเข้าข้อมูลเรียบร้อยแล้ว', 'success');
    refreshAll();
  }

  function mapRecordsForExport(type) {
    var schema = getSchema(type);
    var data = loadRecords(type);
    return data.map(function (r) {
      var obj = {};
      schema.forEach(function (f) {
        if (f.type === 'status') {
          Object.keys(f.statusColumns).forEach(function (label) {
            obj[f.statusColumns[label]] = (r.status === label) ? '✓' : '';
          });
        } else if (f.type === 'boolean') {
          obj[f.label] = r[f.key] ? '✓' : '';
        } else {
          obj[f.label] = (r[f.key] === undefined || r[f.key] === null) ? '' : r[f.key];
        }
      });
      return obj;
    });
  }

  function handleExport() {
    var pkData = mapRecordsForExport('PK');
    var rmData = mapRecordsForExport('RM');
    if (pkData.length === 0 && rmData.length === 0) {
      showToast('ยังไม่มีข้อมูลให้ส่งออก', 'warning');
      return;
    }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pkData.length ? pkData : [{}]), 'PK');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rmData.length ? rmData : [{}]), 'RM');
    XLSX.writeFile(wb, 'QC_Dashboard_export_' + todayStr() + '.xlsx');
    showToast('ดาวน์โหลดไฟล์ Excel เรียบร้อยแล้ว', 'success');
  }

  function handleResetAll() {
    askConfirm('ต้องการลบข้อมูล PK และ RM ทั้งหมดในเบราว์เซอร์นี้ใช่หรือไม่? การลบไม่สามารถย้อนกลับได้', function () {
      localStorage.removeItem(STORAGE_KEYS.PK);
      localStorage.removeItem(STORAGE_KEYS.RM);
      showToast('ล้างข้อมูลทั้งหมดเรียบร้อยแล้ว', 'danger');
      refreshAll();
    });
  }

  function refreshAll() {
    renderDashboard();
    renderTableBody();
    if (activePage === 'daily') renderDailySummary();
    if (activePage === 'analysis') renderAnalysis();
  }

  function tickClock() {
    var el = document.getElementById('clock');
    if (!el) return;
    el.textContent = new Date().toLocaleString('th-TH', { dateStyle: 'medium', timeStyle: 'short' });
  }

  var THEME_KEY = 'qc_theme_mode';
  var ACCENT_KEY = 'qc_theme_accent';

  function applyTheme(mode) {
    document.documentElement.setAttribute('data-theme', mode);
    document.documentElement.setAttribute('data-bs-theme', mode);
    localStorage.setItem(THEME_KEY, mode);
  }

  function applyAccent(accent) {
    document.documentElement.setAttribute('data-accent', accent);
    localStorage.setItem(ACCENT_KEY, accent);
    document.querySelectorAll('.accent-swatch').forEach(function (el) {
      el.classList.toggle('active', el.dataset.accent === accent);
    });
  }

  function initThemeSwitcher() {
    var savedTheme = localStorage.getItem(THEME_KEY) || 'light';
    var savedAccent = localStorage.getItem(ACCENT_KEY) || 'indigo';

    var modeRadio = document.getElementById(savedTheme === 'dark' ? 'modeDark' : 'modeLight');
    if (modeRadio) modeRadio.checked = true;
    applyAccent(savedAccent);

    document.querySelectorAll('input[name="themeMode"]').forEach(function (radio) {
      radio.addEventListener('change', function () {
        if (this.checked) { applyTheme(this.value); renderDashboard(); }
      });
    });
    document.querySelectorAll('.accent-swatch').forEach(function (btn) {
      btn.addEventListener('click', function () { applyAccent(btn.dataset.accent); });
    });
  }

  document.addEventListener('DOMContentLoaded', function () {

    initThemeSwitcher();

    document.querySelectorAll('.sidebar-link').forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        gotoPage(a.dataset.page);
      });
    });
    document.querySelectorAll('[data-goto]').forEach(function (btn) {
      btn.addEventListener('click', function () { gotoPage(btn.dataset.goto); });
    });

    document.getElementById('entryTypePK').addEventListener('change', function () {
      document.getElementById('formPK').classList.remove('d-none');
      document.getElementById('formRM').classList.add('d-none');
    });
    document.getElementById('entryTypeRM').addEventListener('change', function () {
      document.getElementById('formRM').classList.remove('d-none');
      document.getElementById('formPK').classList.add('d-none');
    });
    document.getElementById('formPK').addEventListener('submit', function (e) { handleEntrySubmit(e, 'PK'); });
    document.getElementById('formRM').addEventListener('submit', function (e) { handleEntrySubmit(e, 'RM'); });

    document.querySelectorAll('#tableTab .nav-link').forEach(function (btn) {
      btn.addEventListener('click', function () {
        document.querySelectorAll('#tableTab .nav-link').forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
        tableState.currentTab = btn.dataset.tab;
        tableState.searchQuery = '';
        tableState.currentPage = 1;
        document.getElementById('searchInput').value = '';
        renderTableHead();
        populateStatusFilterOptions();
        renderTableBody();
      });
    });
    document.getElementById('searchInput').addEventListener('input', debounce(function (e) {
      tableState.searchQuery = e.target.value;
      tableState.currentPage = 1;
      renderTableBody();
    }, 150));
    document.getElementById('statusFilter').addEventListener('change', function (e) {
      tableState.statusFilter = e.target.value;
      tableState.currentPage = 1;
      renderTableBody();
    });
    document.getElementById('clearFilterBtn').addEventListener('click', function () {
      document.getElementById('searchInput').value = '';
      document.getElementById('statusFilter').value = '';
      tableState.searchQuery = '';
      tableState.statusFilter = '';
      tableState.currentPage = 1;
      renderTableBody();
    });
    document.getElementById('prevPageBtn').addEventListener('click', function () {
      tableState.currentPage--;
      renderTableBody();
    });
    document.getElementById('nextPageBtn').addEventListener('click', function () {
      tableState.currentPage++;
      renderTableBody();
    });
    document.getElementById('tableBody').addEventListener('click', function (e) {
      var editBtn = e.target.closest('.btn-edit');
      var delBtn = e.target.closest('.btn-delete');
      if (editBtn) openEditModal(tableState.currentTab, editBtn.dataset.id);
      if (delBtn) deleteRecord(tableState.currentTab, delBtn.dataset.id);
    });

    document.querySelectorAll('#analysisTab .nav-link').forEach(function (btn) {
      btn.addEventListener('click', function () {
        document.querySelectorAll('#analysisTab .nav-link').forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
        analysisState.tab = btn.dataset.analysisTab;
        document.getElementById('analysisPanelDefect').classList.toggle('d-none', analysisState.tab !== 'defect');
        document.getElementById('analysisPanelSupplier').classList.toggle('d-none', analysisState.tab !== 'supplier');
        renderAnalysis();
      });
    });
    document.getElementById('analysisTypePK').addEventListener('change', function () {
      if (this.checked) { analysisState.type = 'PK'; renderAnalysis(); }
    });
    document.getElementById('analysisTypeRM').addEventListener('change', function () {
      if (this.checked) { analysisState.type = 'RM'; renderAnalysis(); }
    });
    document.querySelectorAll('#supplierTable th.sortable').forEach(function (th) {
      th.addEventListener('click', function () {
        var field = th.dataset.sort;
        if (analysisState.supplierSort.field === field) {
          analysisState.supplierSort.dir = (analysisState.supplierSort.dir === 'asc') ? 'desc' : 'asc';
        } else {
          analysisState.supplierSort.field = field;
          analysisState.supplierSort.dir = (field === 'passRate') ? 'asc' : 'desc';
        }
        renderSupplierPanel();
      });
    });

    document.getElementById('importFile').addEventListener('change', function () {
      document.getElementById('previewImportBtn').disabled = !this.files.length;
      document.getElementById('importPreview').classList.add('d-none');
    });
    document.getElementById('previewImportBtn').addEventListener('click', handlePreviewImport);
    document.getElementById('confirmImportBtn').addEventListener('click', handleConfirmImport);
    document.getElementById('exportBtn').addEventListener('click', handleExport);
    document.getElementById('resetAllBtn').addEventListener('click', handleResetAll);

    renderTableHead();
    populateStatusFilterOptions();
    initDailySummaryControls();
    var initialPage = (location.hash || '').replace('#', '') || 'dashboard';
    gotoPage(initialPage);

    tickClock();
    setInterval(tickClock, 30000);
  });

})();
