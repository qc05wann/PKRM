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
        inputHtml = '<input type="date" class="form-control" name="' + f.key +
