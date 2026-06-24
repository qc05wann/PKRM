/* =====================================================================
   QC Dashboard — app.js
   เก็บข้อมูลทั้งหมดใน localStorage ของเบราว์เซอร์ (ไม่มีเซิร์ฟเวอร์ภายนอก)
   โครงสร้างข้อมูลอ้างอิงตามไฟล์ Excel งาน QC Dashboard (PK / RM)
   ===================================================================== */
(function () {
  'use strict';

  /* ------------------------------------------------------------------
   * 1) SCHEMA — โครงสร้างคอลัมน์ของ PK และ RM
   *    key       = ชื่อฟิลด์ที่ใช้ภายในระบบ
   *    label     = หัวคอลัมน์ที่แสดงผล / หัวคอลัมน์ตอน export
   *    type      = text | date | number | status
   *    aliases   = ชื่อหัวคอลัมน์ที่เป็นไปได้ตอน "นำเข้า Excel"
   *                (ใส่ไว้หลายแบบเผื่อไฟล์จริงเขียนต่างไปเล็กน้อย)
   * ------------------------------------------------------------------ */
  var PK_SCHEMA = [
    { key: 'receiveNo',   label: 'เลขที่รับเข้า',     type: 'text',   aliases: ['เลขที่รับเข้า', 'receive_no', 'receiveno', 'เลขที่รับ', 'received_no'] },
    { key: 'receiveDate', label: 'วันที่รับเข้า',      type: 'date',   aliases: ['วันที่รับเข้า', 'receive_date', 'receivedate'] },
    { key: 'inspectDate', label: 'วันที่ตรวจสอบ',     type: 'date',   aliases: ['วันที่ตรวจสอบ', 'inspect_date', 'inspectdate'] },
    { key: 'code',        label: 'รหัส',              type: 'text',   aliases: ['รหัส', 'code', 'item_code'] },
    { key: 'name',        label: 'ชื่อบรรจุภัณฑ์',    type: 'text',   aliases: ['ชื่อบรรจุภัณฑ์', 'name', 'item_name', 'ชื่อสินค้า'] },
    { key: 'supplier',    label: 'Supplier',          type: 'text',   aliases: ['supplier', 'ผู้ขาย'] },
    { key: 'lot',         label: 'Lot',               type: 'text',   aliases: ['lot', 'lotno', 'lot_no'] },
    { key: 'qty',         label: 'จำนวน(ชิ้น)',       type: 'number', aliases: ['จำนวน(ชิ้น)', 'จำนวน', 'qty', 'quantity'] },
    { key: 'status',      label: 'สถานะ',             type: 'status', aliases: ['สถานะ', 'status'] },
    { key: 'inspector',   label: 'ผู้ตรวจ',           type: 'text',   aliases: ['ผู้ตรวจ', 'inspector'] },
    { key: 'approver',    label: 'ผู้อนุมัติ',        type: 'text',   aliases: ['ผู้อนุมัติ', 'approver'] },
    { key: 'defectType',  label: 'ประเภท Defect',     type: 'text',   aliases: ['ประเภทdefect', 'defect', 'defect_type', 'defecttype'] }
  ];

  var RM_SCHEMA = [
    { key: 'receiveNo',   label: 'Received_no',       type: 'text',   aliases: ['received_no', 'receivedno', 'เลขที่รับเข้า'] },
    { key: 'receiveDate', label: 'วันที่รับเข้า',      type: 'date',   aliases: ['วันที่รับเข้า', 'receive_date', 'receivedate'] },
    { key: 'code',        label: 'รหัสวัตถุดิบ',      type: 'text',   aliases: ['รหัสวัตถุดิบ', 'code', 'rm_code'] },
    { key: 'name',        label: 'ชื่อวัตถุดิบ',      type: 'text',   aliases: ['ชื่อวัตถุดิบ', 'name', 'rm_name'] },
    { key: 'supplier',    label: 'Supplier',          type: 'text',   aliases: ['supplier', 'ผู้ขาย'] },
    { key: 'lot',         label: 'Lot',               type: 'text',   aliases: ['lot', 'lotno', 'lot_no'] },
    { key: 'status',      label: 'สถานะอนุมัติ',      type: 'status', aliases: ['สถานะอนุมัติ', 'status', 'approval_status'] },
    { key: 'inspector',   label: 'ผู้ตรวจ',           type: 'text',   aliases: ['ผู้ตรวจ', 'inspector'] },
    { key: 'approver',    label: 'ผู้อนุมัติ',        type: 'text',   aliases: ['ผู้อนุมัติ', 'approver'] }
  ];

  var PK_STATUSES = ['ผ่าน', 'Quarantine', 'Rejected'];
  var RM_STATUSES = ['Approved', 'Quarantine', 'Rejected', 'Retain'];

  var STATUS_COLOR_HEX = { pass: '#1B998B', warn: '#E8A33D', danger: '#D7263D', retain: '#5E60CE', neutral: '#5C6784' };

  var STORAGE_KEYS = { PK: 'qc_pk_records', RM: 'qc_rm_records' };

  /* ------------------------------------------------------------------
   * 2) STORAGE HELPERS
   * ------------------------------------------------------------------ */
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

  /* ------------------------------------------------------------------
   * 3) UTILITIES
   * ------------------------------------------------------------------ */
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
        'Rejected':   { cls: 'status-danger', icon: 'bi-x-circle-fill',           color: 'danger' },
        'Retain':     { cls: 'status-retain', icon: 'bi-archive-fill',            color: 'retain' }
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

  /* ------------------------------------------------------------------
   * 4) ROUTER / NAVIGATION
   * ------------------------------------------------------------------ */
  var PAGES = ['dashboard', 'entry', 'table', 'io'];
  var TITLES = {
    dashboard: 'ภาพรวม (Dashboard)',
    entry: 'บันทึกข้อมูล',
    table: 'ตารางข้อมูล / ค้นหา',
    io: 'นำเข้า / ส่งออก Excel'
  };

  function gotoPage(page) {
    if (PAGES.indexOf(page) === -1) page = 'dashboard';
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

    // ปิดเมนูข้าง (offcanvas) บนมือถือ/แท็บเล็ตหลังกดเมนู
    var sidebarEl = document.getElementById('sidebar');
    var oc = bootstrap.Offcanvas.getInstance(sidebarEl);
    if (oc) oc.hide();
  }

  /* ------------------------------------------------------------------
   * 5) DASHBOARD
   * ------------------------------------------------------------------ */
  var chartInstances = { PK: null, RM: null };

  function computeKPI(type, data) {
    var statuses = getStatusList(type);
    var counts = {};
    statuses.forEach(function (s) { counts[s] = 0; });
    data.forEach(function (r) { if (counts.hasOwnProperty(r.status)) counts[r.status]++; });
    var total = data.length;
    var passKey = type === 'PK' ? 'ผ่าน' : 'Approved';
    var passRate = total ? ((counts[passKey] / total) * 100).toFixed(1) : '0.0';
    return { total: total, counts: counts, passRate: passRate };
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
    chartInstances[type] = new Chart(canvas.getContext('2d'), {
      type: 'doughnut',
      data: { labels: statuses, datasets: [{ data: values, backgroundColor: colors, borderWidth: 0 }] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '62%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 10, font: { family: 'Sarabun', size: 11 } } }
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

  /* ------------------------------------------------------------------
   * 6) DATA ENTRY
   * ------------------------------------------------------------------ */
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

  /* ------------------------------------------------------------------
   * 7) DATA TABLE / SEARCH / EDIT / DELETE
   * ------------------------------------------------------------------ */
  var tableState = { currentTab: 'PK', searchQuery: '', statusFilter: '' };

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
      return;
    }
    tableEl.classList.remove('d-none');
    emptyEl.classList.add('d-none');

    tbody.innerHTML = filtered.slice().reverse().map(function (r) {
      var cells = schema.map(function (f) {
        var val = r[f.key];
        if (f.type === 'status') return '<td>' + statusBadgeHTML(type, val) + '</td>';
        if (f.type === 'date') return '<td>' + formatDateDisplay(val) + '</td>';
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

  /* ------------------------------------------------------------------
   * 8) IMPORT (Excel -> localStorage)
   * ------------------------------------------------------------------ */
  var pendingImport = { PK: [], RM: [] };

  function normalizeHeader(h) {
    return String(h).trim().toLowerCase().replace(/[\s_()]/g, '');
  }

  function dateToInputValue(d) {
    var tz = d.getTimezoneOffset();
    var local = new Date(d.getTime() - tz * 60000);
    return local.toISOString().slice(0, 10);
  }

  function normalizeDateValue(val) {
    if (!val) return '';
    if (val instanceof Date && !isNaN(val)) return dateToInputValue(val);
    var s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    return s;
  }

  function mapRowToRecord(type, rawRow) {
    var schema = getSchema(type);
    var normalizedRow = {};
    Object.keys(rawRow).forEach(function (k) { normalizedRow[normalizeHeader(k)] = rawRow[k]; });

    var record = {};
    schema.forEach(function (f) {
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
      else val = (typeof val === 'string') ? val.trim() : (val === undefined ? '' : val);
      record[f.key] = val;
    });
    record.id = genId();
    record.createdAt = Date.now();
    return record;
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
        var workbook = XLSX.read(data, { type: 'array', cellDates: true });

        var pkSheetName = findSheetName(workbook, ['pk', 'บรรจุภัณฑ์']);
        var rmSheetName = findSheetName(workbook, ['rm', 'วัตถุดิบ']);

        pendingImport.PK = [];
        pendingImport.RM = [];
        var logLines = [];

        if (pkSheetName) {
          var pkRows = XLSX.utils.sheet_to_json(workbook.Sheets[pkSheetName], { defval: '' });
          pendingImport.PK = pkRows.map(function (row) { return mapRowToRecord('PK', row); });
          logLines.push('พบชีต "' + pkSheetName + '" → ' + pendingImport.PK.length + ' แถว (PK)');
        } else {
          logLines.push('ไม่พบชีตที่ตรงกับข้อมูล PK ในไฟล์นี้');
        }

        if (rmSheetName) {
          var rmRows = XLSX.utils.sheet_to_json(workbook.Sheets[rmSheetName], { defval: '' });
          pendingImport.RM = rmRows.map(function (row) { return mapRowToRecord('RM', row); });
          logLines.push('พบชีต "' + rmSheetName + '" → ' + pendingImport.RM.length + ' แถว (RM)');
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

    if (pendingImport.PK.length) {
      var pkArr = mode === 'replace' ? pendingImport.PK : loadRecords('PK').concat(pendingImport.PK);
      saveRecords('PK', pkArr);
    } else if (mode === 'replace') {
      saveRecords('PK', []);
    }

    if (pendingImport.RM.length) {
      var rmArr = mode === 'replace' ? pendingImport.RM : loadRecords('RM').concat(pendingImport.RM);
      saveRecords('RM', rmArr);
    } else if (mode === 'replace') {
      saveRecords('RM', []);
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

    showToast('นำเข้าข้อมูลเรียบร้อยแล้ว', 'success');
    refreshAll();
  }

  /* ------------------------------------------------------------------
   * 9) EXPORT (localStorage -> Excel)
   * ------------------------------------------------------------------ */
  function mapRecordsForExport(type) {
    var schema = getSchema(type);
    var data = loadRecords(type);
    return data.map(function (r) {
      var obj = {};
      schema.forEach(function (f) { obj[f.label] = (r[f.key] === undefined || r[f.key] === null) ? '' : r[f.key]; });
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

  /* ------------------------------------------------------------------
   * 10) REFRESH ALL (เรียกทุกครั้งที่ข้อมูลเปลี่ยน)
   * ------------------------------------------------------------------ */
  function refreshAll() {
    renderDashboard();
    renderTableBody();
  }

  /* ------------------------------------------------------------------
   * 11) CLOCK
   * ------------------------------------------------------------------ */
  function tickClock() {
    var el = document.getElementById('clock');
    if (!el) return;
    el.textContent = new Date().toLocaleString('th-TH', { dateStyle: 'medium', timeStyle: 'short' });
  }

  /* ------------------------------------------------------------------
   * 12) INIT
   * ------------------------------------------------------------------ */
  document.addEventListener('DOMContentLoaded', function () {

    // --- Sidebar navigation ---
    document.querySelectorAll('.sidebar-link').forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        gotoPage(a.dataset.page);
      });
    });
    document.querySelectorAll('[data-goto]').forEach(function (btn) {
      btn.addEventListener('click', function () { gotoPage(btn.dataset.goto); });
    });

    // --- Entry: type switch ---
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

    // --- Table: tabs / search / filter ---
    document.querySelectorAll('#tableTab .nav-link').forEach(function (btn) {
      btn.addEventListener('click', function () {
        document.querySelectorAll('#tableTab .nav-link').forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
        tableState.currentTab = btn.dataset.tab;
        tableState.searchQuery = '';
        document.getElementById('searchInput').value = '';
        renderTableHead();
        populateStatusFilterOptions();
        renderTableBody();
      });
    });
    document.getElementById('searchInput').addEventListener('input', debounce(function (e) {
      tableState.searchQuery = e.target.value;
      renderTableBody();
    }, 150));
    document.getElementById('statusFilter').addEventListener('change', function (e) {
      tableState.statusFilter = e.target.value;
      renderTableBody();
    });
    document.getElementById('clearFilterBtn').addEventListener('click', function () {
      document.getElementById('searchInput').value = '';
      document.getElementById('statusFilter').value = '';
      tableState.searchQuery = '';
      tableState.statusFilter = '';
      renderTableBody();
    });
    document.getElementById('tableBody').addEventListener('click', function (e) {
      var editBtn = e.target.closest('.btn-edit');
      var delBtn = e.target.closest('.btn-delete');
      if (editBtn) openEditModal(tableState.currentTab, editBtn.dataset.id);
      if (delBtn) deleteRecord(tableState.currentTab, delBtn.dataset.id);
    });

    // --- Import / Export ---
    document.getElementById('importFile').addEventListener('change', function () {
      document.getElementById('previewImportBtn').disabled = !this.files.length;
      document.getElementById('importPreview').classList.add('d-none');
    });
    document.getElementById('previewImportBtn').addEventListener('click', handlePreviewImport);
    document.getElementById('confirmImportBtn').addEventListener('click', handleConfirmImport);
    document.getElementById('exportBtn').addEventListener('click', handleExport);
    document.getElementById('resetAllBtn').addEventListener('click', handleResetAll);

    // --- Initial render ---
    renderTableHead();
    populateStatusFilterOptions();
    var initialPage = (location.hash || '').replace('#', '') || 'dashboard';
    gotoPage(initialPage);

    tickClock();
    setInterval(tickClock, 30000);
  });

})();
