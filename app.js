/* ============================================================================
 * Gestión de Equipos en Servicio Técnico
 * Aplicación de registro por etapas con persistencia local (localStorage)
 * y exportación a Excel (.xlsx).
 *
 * Cada etapa del proceso se registra como un registro independiente: se crea
 * por sí solo, sin obligar a seguir la secuencia. El folio de la solicitud de
 * trabajo se ingresa manualmente; en todos los casos los técnicos se eligen
 * desde una lista desplegable.
 * ========================================================================== */
(function () {
  'use strict';

  // ---------------------------------------------------------------- Constantes
  var STORAGE_KEY = 'gec_v1';

  var TECNICOS_DEFAULT = [
    'Carlos Bahamondes Seguel',
    'Cristián Beltrán Oviedo',
    'Cristina Rozas Urrutia',
    'Daniel Díaz Neira',
    'Ignacio Berner Bergara',
    'Macarena Toledo',
    'Marco Ulloa',
    'Matías Soazo Garrido',
    'Ricardo Matus Aroca',
    'Tito Millapán Riquelme'
  ];
  var SUPERVISOR = 'Cristián Beltrán Oviedo';
  var EMPRESAS_DEFAULT = [];

  // Definición de las etapas (esquema que dirige formularios, tablas y export).
  // tipos de campo: folio_manual | folio_ref | fecha | tecnico | empresa |
  //                 select | equipo | text | numero | textarea
  var ETAPAS = [
    {
      id: 'solicitud', nombre: 'Solicitud de trabajo', icono: '📝',
      grupo: 'Solicitud', via: 'Inicio',
      desc: 'Etapa común de inicio. El folio se ingresa manualmente. Luego el flujo se deriva a la Vía A o a la Vía B.',
      campos: [
        { key: 'folio', label: 'Folio (manual)', tipo: 'folio_manual', req: true, ancho: 14, hint: 'Identificador de la solicitud. Se ingresa manualmente.' },
        { key: 'fecha', label: 'Fecha de solicitud', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'tecnico', label: 'Técnico asignado', tipo: 'tecnico', req: true, ancho: 24 },
        { key: 'via', label: 'Vía', tipo: 'select', opciones: ['Vía A — Servicio técnico', 'Vía B — Visita técnica / diagnóstico'], ancho: 26 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'envio', nombre: 'Envío a servicio técnico', icono: '📤',
      grupo: 'Vía A · Servicio técnico', via: 'Vía A',
      desc: 'Vía A · 6.1 — Salida del equipo a la empresa de servicio técnico. Se cursa sobre solicitudes vigentes / no cerradas.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14, hint: 'Folio de la solicitud asociada (preferentemente vigente / no cerrada).' },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha del envío', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'numero_envio', label: 'Número de envío', tipo: 'text', ancho: 16 },
        { key: 'tecnico', label: 'Técnico asignado', tipo: 'tecnico', req: true, ancho: 24 },
        { key: 'empresa', label: 'Empresa de servicio técnico', tipo: 'empresa', ancho: 24 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'estado_st', nombre: 'Estado en servicio técnico', icono: '🛠️',
      grupo: 'Vía A · Servicio técnico', via: 'Vía A',
      desc: 'Vía A · 6.2 — Seguimiento intermedio mientras el equipo permanece en el taller.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de actualización', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'estado', label: 'Estado', tipo: 'select', req: true, opciones: ['En diagnóstico', 'En reparación', 'Reparado', 'Sin solución'], ancho: 18 },
        { key: 'tecnico', label: 'Técnico', tipo: 'tecnico', ancho: 24 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'recepcion', nombre: 'Recepción del equipo', icono: '📥',
      grupo: 'Vía A · Servicio técnico', via: 'Vía A',
      desc: 'Vía A · 6.3 — Ingreso del equipo de vuelta y registro de su estado. Exclusiva de la Vía A.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de recepción', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'numero_guia', label: 'N° de guía de despacho', tipo: 'text', ancho: 18 },
        { key: 'tecnico', label: 'Técnico (recibe)', tipo: 'tecnico', ancho: 24 },
        { key: 'estado_equipo', label: 'Estado del equipo', tipo: 'select', opciones: ['Operativo', 'No operativo'], ancho: 16 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'diagnostico', nombre: 'Visita técnica / diagnóstico', icono: '🔍',
      grupo: 'Vía B · En sitio', via: 'Vía B',
      desc: 'Vía B — El equipo permanece en sitio. No hay envío ni recepción. Se registra el estado resultante del diagnóstico.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de diagnóstico', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'tecnico', label: 'Técnico', tipo: 'tecnico', req: true, ancho: 24 },
        { key: 'estado_equipo', label: 'Estado del equipo', tipo: 'select', opciones: ['No operativo', 'Operativo'], def: 'No operativo', ancho: 16 },
        { key: 'observaciones', label: 'Resultado / observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'cotizacion', nombre: 'Solicitud de cotización', icono: '💬',
      grupo: 'Subflujo comercial', via: 'Comercial',
      desc: 'Subflujo comercial · 8.1 — Opcional. Solo cuando la reparación requiere adquisición de bienes o servicios.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de cotización', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'numero_cotizacion', label: 'N° de cotización', tipo: 'text', ancho: 16 },
        { key: 'empresa', label: 'Empresa (proveedor)', tipo: 'empresa', ancho: 24 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'gestion_oc', nombre: 'Gestión de orden de compra', icono: '📑',
      grupo: 'Subflujo comercial', via: 'Comercial',
      desc: 'Subflujo comercial · 8.2 — Dos caminos excluyentes: Trato directo (informe) o Compra ágil.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'tipo_compra', label: 'Tipo de compra', tipo: 'select', req: true, opciones: ['Trato directo', 'Compra ágil'], ancho: 16 },
        { key: 'fecha', label: 'Fecha (informe / compra)', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'numero_informe', label: 'N° de informe (trato directo)', tipo: 'text', ancho: 20, hint: 'Solo aplica en Trato directo.' },
        { key: 'tecnico', label: 'Técnico', tipo: 'tecnico', ancho: 24 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'emision_oc', nombre: 'Emisión de orden de compra', icono: '🧾',
      grupo: 'Subflujo comercial', via: 'Comercial',
      desc: 'Subflujo comercial · 8.3 — Se emite tras la aprobación del presupuesto.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de OC', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'numero_oc', label: 'N° de OC', tipo: 'text', ancho: 16 },
        { key: 'tecnico', label: 'Técnico', tipo: 'tecnico', ancho: 24 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'reparacion', nombre: 'Reparación del equipo', icono: '🔧',
      grupo: 'Cierre del ciclo', via: 'Común',
      desc: 'Sección 9 — Intervención sobre el equipo, común a ambas vías.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de reparación', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'tecnico', label: 'Técnico', tipo: 'tecnico', req: true, ancho: 24 },
        { key: 'resultado', label: 'Resultado', tipo: 'select', req: true, opciones: ['Operativo', 'No operativo'], ancho: 16 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    },
    {
      id: 'cierre', nombre: 'Cierre del ciclo', icono: '✅',
      grupo: 'Cierre del ciclo', via: 'Cierre',
      desc: 'Sección 10 — El equipo queda Operativo y la solicitud se marca como cerrada.',
      campos: [
        { key: 'folio', label: 'Solicitud de trabajo (folio)', tipo: 'folio_ref', ancho: 14 },
        { key: 'equipo', label: 'Equipo (listado crítico)', tipo: 'equipo', col: 'full' },
        { key: 'fecha', label: 'Fecha de cierre', tipo: 'fecha', req: true, ancho: 14 },
        { key: 'estado_final', label: 'Estado final', tipo: 'select', opciones: ['Operativo', 'No operativo'], def: 'Operativo', ancho: 16 },
        { key: 'tecnico', label: 'Técnico', tipo: 'tecnico', ancho: 24 },
        { key: 'observaciones', label: 'Observaciones', tipo: 'textarea', col: 'full', ancho: 40 }
      ]
    }
  ];

  var ETAPAS_BY_ID = {};
  ETAPAS.forEach(function (e, i) { e._orden = i; ETAPAS_BY_ID[e.id] = e; });

  var GRUPOS_NAV = [
    { label: 'Inicio', items: ['__dashboard'] },
    { label: 'Solicitud', items: ['solicitud'] },
    { label: 'Vía A · Servicio técnico', items: ['envio', 'estado_st', 'recepcion'] },
    { label: 'Vía B · En sitio', items: ['diagnostico'] },
    { label: 'Subflujo comercial', items: ['cotizacion', 'gestion_oc', 'emision_oc'] },
    { label: 'Cierre del ciclo', items: ['reparacion', 'cierre'] },
    { label: 'Gestión', items: ['__todos', '__config'] }
  ];

  var EQUIPOS = window.EQUIPOS || [];

  // ----------------------------------------------------------------- Estado/DB
  var DB = cargarDB();
  var STATE = { view: '__dashboard', editId: null };

  function cargarDB() {
    var db = null;
    try { db = JSON.parse(localStorage.getItem(STORAGE_KEY)); } catch (e) { db = null; }
    if (!db || typeof db !== 'object') db = {};
    if (!db.registros) db.registros = {};
    if (!db.config) db.config = {};
    if (!Array.isArray(db.config.tecnicos) || !db.config.tecnicos.length) db.config.tecnicos = TECNICOS_DEFAULT.slice();
    if (!Array.isArray(db.config.empresas)) db.config.empresas = EMPRESAS_DEFAULT.slice();
    ETAPAS.forEach(function (e) { if (!Array.isArray(db.registros[e.id])) db.registros[e.id] = []; });
    return db;
  }
  function guardarDB() { localStorage.setItem(STORAGE_KEY, JSON.stringify(DB)); }

  // ------------------------------------------------------------------- Helpers
  function el(tag, attrs, children) {
    var n = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (attrs[k] == null) continue;
      if (k === 'class') n.className = attrs[k];
      else if (k === 'html') n.innerHTML = attrs[k];
      else n.setAttribute(k, attrs[k]);
    }
    if (children != null) {
      (Array.isArray(children) ? children : [children]).forEach(function (c) {
        if (c == null || c === false) return;
        n.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
      });
    }
    return n;
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }
  function hoyISO() {
    var d = new Date(), z = function (n) { return String(n).padStart(2, '0'); };
    return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate());
  }
  function uid() { return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function cmpNat(a, b) {
    return String(a == null ? '' : a).localeCompare(String(b == null ? '' : b), 'es', { numeric: true, sensitivity: 'base' });
  }
  function fmtFechaHora(iso) {
    if (!iso) return '';
    var d = new Date(iso); if (isNaN(d)) return '';
    var z = function (n) { return String(n).padStart(2, '0'); };
    return z(d.getDate()) + '-' + z(d.getMonth() + 1) + '-' + d.getFullYear() + ' ' + z(d.getHours()) + ':' + z(d.getMinutes());
  }
  function fmtFecha(iso) {
    if (!iso) return '';
    var p = String(iso).split('-');
    return p.length === 3 ? (p[2] + '-' + p[1] + '-' + p[0]) : iso;
  }
  function equipoCorto(eq) {
    if (!eq) return '';
    return (eq.inv || '(s/inv)') + ' — ' + (eq.nombre || '');
  }

  // -------------------------------------------------------------- Búsqueda eq.
  function buscarEquipos(q, limite) {
    var tokens = q.toLowerCase().split(/\s+/).filter(Boolean);
    var out = [];
    for (var i = 0; i < EQUIPOS.length; i++) {
      var e = EQUIPOS[i];
      var hay = ((e.inventario || '') + ' ' + (e.equipo || '') + ' ' + (e.serie || '') + ' ' +
        (e.marca || '') + ' ' + (e.modelo || '') + ' ' + (e.servicio || '') + ' ' +
        (e.unidad || '') + ' ' + (e.ubicacion || '')).toLowerCase();
      var ok = true;
      for (var t = 0; t < tokens.length; t++) { if (hay.indexOf(tokens[t]) < 0) { ok = false; break; } }
      if (ok) { out.push(e); if (out.length >= limite) break; }
    }
    return out;
  }

  // ------------------------------------------------------------------- Toasts
  function toast(msg, tipo) {
    var cont = document.getElementById('toasts');
    var t = el('div', { class: 'toast ' + (tipo || '') }, msg);
    cont.appendChild(t);
    setTimeout(function () { t.style.opacity = '0'; t.style.transition = 'opacity .3s'; }, 2600);
    setTimeout(function () { if (t.parentNode) t.parentNode.removeChild(t); }, 3000);
  }

  // ------------------------------------------------------------- Selector eq.
  function buildEquipoPicker(inicial) {
    var wrap = el('div', { class: 'equipo-pick' });
    var input = el('input', { type: 'text', autocomplete: 'off', placeholder: 'Buscar por inventario, equipo, serie, marca, servicio…' });
    var results = el('div', { class: 'equipo-results' });
    var chip = el('div', { class: 'equipo-chip' });
    var selected = inicial || null;

    function pintarChip() {
      if (selected) {
        chip.innerHTML = '<span class="x" title="Quitar">✕</span><strong>' + esc(selected.inv || '(sin inventario)') + '</strong> — ' +
          esc(selected.nombre || '') + ' <span style="color:#6b7780">· ' + esc(selected.servicio || '') +
          (selected.serie ? (' · Serie ' + esc(selected.serie)) : '') +
          (selected.ubicacion ? (' · ' + esc(selected.ubicacion)) : '') + '</span>';
        chip.classList.add('show');
        input.style.display = 'none';
        chip.querySelector('.x').onclick = function () { selected = null; pintarChip(); };
      } else {
        chip.classList.remove('show');
        input.style.display = '';
      }
    }
    input.addEventListener('input', function () {
      var q = input.value.trim();
      results.innerHTML = '';
      if (!q) { results.classList.remove('show'); return; }
      var ms = buscarEquipos(q, 40);
      if (!ms.length) { results.innerHTML = '<div class="empty">Sin coincidencias en el listado crítico</div>'; results.classList.add('show'); return; }
      ms.forEach(function (m) {
        var it = el('div', { class: 'item' });
        it.innerHTML = '<div class="t">' + esc(m.inventario || '(sin inventario)') + ' — ' + esc(m.equipo) + '</div>' +
          '<div class="m">' + esc(m.servicio || '') + ' · ' + esc([m.marca, m.modelo].filter(Boolean).join(' ')) +
          (m.serie ? (' · Serie ' + esc(m.serie)) : '') + (m.ubicacion ? (' · ' + esc(m.ubicacion)) : '') + '</div>';
        it.onclick = function () {
          selected = { inv: m.inventario, nombre: m.equipo, servicio: m.servicio, serie: m.serie, marca: m.marca, modelo: m.modelo, unidad: m.unidad, ubicacion: m.ubicacion };
          input.value = ''; results.classList.remove('show'); pintarChip();
        };
        results.appendChild(it);
      });
      results.classList.add('show');
    });
    document.addEventListener('click', function (e) { if (!wrap.contains(e.target)) results.classList.remove('show'); });

    wrap.appendChild(input); wrap.appendChild(results); wrap.appendChild(chip);
    pintarChip();
    return { wrap: wrap, get: function () { return selected; } };
  }

  // ----------------------------------------------------------- Form genérico
  function buildForm(etapa, record) {
    var grid = el('div', { class: 'form-grid' });
    var controls = {};

    etapa.campos.forEach(function (campo) {
      var clazz = 'field' + (campo.col === 'full' ? ' col-full' : (campo.col === '2' ? ' col-2' : ''));
      var field = el('div', { class: clazz });
      var lbl = el('label', {}, campo.label + ' ');
      if (campo.req) lbl.appendChild(el('span', { class: 'req' }, '*'));
      field.appendChild(lbl);

      var ctrl;
      if (campo.tipo === 'equipo') {
        var picker = buildEquipoPicker(record ? record[campo.key] : null);
        field.appendChild(picker.wrap);
        controls[campo.key] = { get: function () { return picker.get(); } };
      } else if (campo.tipo === 'tecnico') {
        ctrl = el('select');
        ctrl.appendChild(el('option', { value: '' }, '— Seleccionar técnico —'));
        DB.config.tecnicos.forEach(function (t) { ctrl.appendChild(el('option', { value: t }, t)); });
        if (record && record[campo.key]) {
          if (DB.config.tecnicos.indexOf(record[campo.key]) < 0) ctrl.appendChild(el('option', { value: record[campo.key] }, record[campo.key]));
          ctrl.value = record[campo.key];
        }
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value; } };
      } else if (campo.tipo === 'select') {
        ctrl = el('select');
        if (!campo.req) ctrl.appendChild(el('option', { value: '' }, '—'));
        campo.opciones.forEach(function (o) { ctrl.appendChild(el('option', { value: o }, o)); });
        var sv = record ? record[campo.key] : campo.def;
        if (sv != null && sv !== '') ctrl.value = sv;
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value; } };
      } else if (campo.tipo === 'empresa') {
        ctrl = el('input', { type: 'text', list: 'dl-empresas', placeholder: 'Nombre de la empresa…' });
        if (record && record[campo.key]) ctrl.value = record[campo.key];
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value.trim(); } };
      } else if (campo.tipo === 'folio_ref') {
        ctrl = el('input', { type: 'text', list: 'dl-folios', placeholder: 'Folio de la solicitud…' });
        if (record && record[campo.key]) ctrl.value = record[campo.key];
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value.trim(); } };
      } else if (campo.tipo === 'folio_manual') {
        ctrl = el('input', { type: 'text', placeholder: 'Ingrese el folio manualmente…' });
        if (record && record[campo.key]) ctrl.value = record[campo.key];
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value.trim(); } };
      } else if (campo.tipo === 'fecha') {
        ctrl = el('input', { type: 'date' });
        ctrl.value = record ? (record[campo.key] || '') : hoyISO();
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value; } };
      } else if (campo.tipo === 'textarea') {
        ctrl = el('textarea', { rows: '2', placeholder: 'Comentarios relevantes…' });
        if (record && record[campo.key]) ctrl.value = record[campo.key];
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value.trim(); } };
      } else {
        ctrl = el('input', { type: campo.tipo === 'numero' ? 'number' : 'text' });
        if (record && record[campo.key] != null) ctrl.value = record[campo.key];
        field.appendChild(ctrl);
        controls[campo.key] = { get: function () { return ctrl.value.trim(); } };
      }
      if (campo.hint) field.appendChild(el('div', { class: 'hint' }, campo.hint));
      grid.appendChild(field);
    });

    return { grid: grid, controls: controls };
  }

  function collectForm(etapa, controls) {
    var rec = {};
    etapa.campos.forEach(function (campo) {
      rec[campo.key] = controls[campo.key].get();
    });
    // Validación de obligatorios
    for (var i = 0; i < etapa.campos.length; i++) {
      var c = etapa.campos[i];
      if (!c.req) continue;
      var v = rec[c.key];
      var vacio = (c.tipo === 'equipo') ? !v : (v == null || v === '');
      if (vacio) throw new Error('Falta completar el campo obligatorio: «' + c.label + '».');
    }
    return rec;
  }

  // --------------------------------------------------------------- Datalists
  function refrescarDatalists() {
    var dlE = document.getElementById('dl-empresas'); dlE.innerHTML = '';
    DB.config.empresas.slice().sort(cmpNat).forEach(function (e) { dlE.appendChild(el('option', { value: e })); });
    var dlF = document.getElementById('dl-folios'); dlF.innerHTML = '';
    folioList().forEach(function (f) { dlF.appendChild(el('option', { value: f })); });
  }
  function folioList() {
    var s = {};
    Object.keys(DB.registros).forEach(function (k) {
      DB.registros[k].forEach(function (r) { if (r.folio) s[r.folio] = 1; });
    });
    return Object.keys(s).sort(cmpNat);
  }

  // ------------------------------------------------------------------ Pills
  function pillFor(key, value) {
    if (!value) return document.createTextNode('');
    var cls = 'pill pill-gray';
    if (key === 'via') cls = /Vía A/.test(value) ? 'pill pill-via-a' : 'pill pill-via-b';
    else if (key === 'resultado' || key === 'estado_equipo' || key === 'estado_final') cls = /No operativo/i.test(value) ? 'pill pill-no' : 'pill pill-ok';
    else if (key === 'estado') cls = /Sin soluci/i.test(value) ? 'pill pill-no' : (/Reparado/i.test(value) ? 'pill pill-ok' : 'pill pill-st');
    else if (key === 'tipo_compra') cls = 'pill pill-via-a';
    return el('span', { class: cls }, value);
  }

  // =========================================================== Render general
  var contentEl, viewTitleEl, viewSubEl;

  function navegar(view) {
    STATE.view = view;
    STATE.editId = null;
    render();
    document.getElementById('sidebar').classList.remove('open');
    document.getElementById('backdrop').classList.remove('show');
    window.scrollTo(0, 0);
  }

  function render() {
    refrescarDatalists();
    renderSidebar();
    if (STATE.view === '__dashboard') renderDashboard();
    else if (STATE.view === '__todos') renderTodos();
    else if (STATE.view === '__config') renderConfig();
    else renderEtapa(STATE.view);
  }

  function renderSidebar() {
    var nav = document.getElementById('nav');
    nav.innerHTML = '';
    GRUPOS_NAV.forEach(function (g) {
      nav.appendChild(el('div', { class: 'group-label' }, g.label));
      g.items.forEach(function (id) {
        var label, icono, badge = null;
        if (id === '__dashboard') { label = 'Resumen'; icono = '📊'; }
        else if (id === '__todos') { label = 'Todos los registros'; icono = '🗂️'; badge = totalRegistros(); }
        else if (id === '__config') { label = 'Configuración'; icono = '⚙️'; }
        else { var et = ETAPAS_BY_ID[id]; label = et.nombre; icono = et.icono; badge = DB.registros[id].length; }
        var a = el('a', { class: STATE.view === id ? 'active' : '' }, [
          el('span', { class: 'ico' }, icono),
          el('span', {}, label),
          (badge != null && badge > 0) ? el('span', { class: 'badge' }, String(badge)) : null
        ]);
        a.onclick = function () { navegar(id); };
        nav.appendChild(a);
      });
    });
  }

  function setTitulo(t, s) { viewTitleEl.textContent = t; viewSubEl.textContent = s || ''; }

  function totalRegistros() {
    var n = 0; ETAPAS.forEach(function (e) { n += DB.registros[e.id].length; }); return n;
  }

  // ------------------------------------------------------------- Dashboard
  function renderDashboard() {
    setTitulo('Resumen del proceso', 'Gestión de equipos en servicio técnico · versión 2.0');
    contentEl.innerHTML = '';

    var folios = folioList();
    var cerrados = {}; DB.registros.cierre.forEach(function (r) { if (r.folio) cerrados[r.folio] = 1; });
    var equiposSet = {};
    ETAPAS.forEach(function (e) { DB.registros[e.id].forEach(function (r) { if (r.equipo && r.equipo.inv) equiposSet[r.equipo.inv] = 1; }); });
    var vigentes = folios.filter(function (f) { return !cerrados[f]; }).length;

    var stats = el('div', { class: 'stat-grid' });
    function stat(n, l, accent) { return el('div', { class: 'stat' + (accent ? ' accent' : '') }, [el('div', { class: 'n' }, String(n)), el('div', { class: 'l' }, l)]); }
    stats.appendChild(stat(totalRegistros(), 'Registros totales'));
    stats.appendChild(stat(folios.length, 'Folios de solicitud'));
    stats.appendChild(stat(vigentes, 'Solicitudes vigentes'));
    stats.appendChild(stat(Object.keys(cerrados).length, 'Ciclos cerrados', true));
    stats.appendChild(stat(Object.keys(equiposSet).length, 'Equipos intervenidos'));
    contentEl.appendChild(stats);

    var banner = el('div', { class: 'banner' },
      'Cada etapa se registra de forma independiente: puede crear cualquier registro sin necesidad de completar las etapas previas. ' +
      'El folio de la solicitud se ingresa manualmente y, en el resto de las etapas, puede reutilizarlo desde la lista.');
    contentEl.appendChild(banner);

    var card = el('div', { class: 'card' });
    card.appendChild(el('div', { class: 'card-head' }, [el('h3', {}, 'Etapas del proceso'), el('span', { class: 'desc' }, 'Haga clic en una etapa para registrar.')]));
    var body = el('div', { class: 'card-body' });
    var flow = el('div', { class: 'flow' });
    ETAPAS.forEach(function (et) {
      var step = el('div', { class: 'step' }, [
        el('div', { class: 'tag' }, et.via),
        el('div', { class: 'name' }, et.icono + ' ' + et.nombre),
        el('div', { class: 'cnt' }, DB.registros[et.id].length + ' registro' + (DB.registros[et.id].length === 1 ? '' : 's'))
      ]);
      step.onclick = function () { navegar(et.id); };
      flow.appendChild(step);
    });
    body.appendChild(flow);
    card.appendChild(body);
    contentEl.appendChild(card);

    // Actividad reciente
    var todos = [];
    ETAPAS.forEach(function (et) { DB.registros[et.id].forEach(function (r) { todos.push({ et: et, r: r }); }); });
    todos.sort(function (a, b) { return (b.r._createdAt || '').localeCompare(a.r._createdAt || ''); });
    var recientes = todos.slice(0, 8);

    var card2 = el('div', { class: 'card' });
    card2.appendChild(el('div', { class: 'card-head' }, [el('h3', {}, 'Actividad reciente')]));
    if (!recientes.length) {
      card2.appendChild(el('div', { class: 'empty-state' }, [el('div', { class: 'big' }, '🗒️'), el('div', {}, 'Aún no hay registros. Comience creando una solicitud de trabajo o cualquier otra etapa.')]));
    } else {
      var wrap = el('div', { class: 'tabla-wrap' });
      var t = el('table', { class: 'data' });
      t.appendChild(el('thead', {}, el('tr', {}, [th('Registrado'), th('Etapa'), th('Folio'), th('Equipo'), th('Técnico'), th('Detalle')])));
      var tb = el('tbody');
      recientes.forEach(function (x) {
        tb.appendChild(el('tr', {}, [
          td(fmtFechaHora(x.r._createdAt)),
          td(el('span', { class: 'tag-etapa' }, x.et.nombre)),
          td(x.r.folio || '—'),
          td(equipoCorto(x.r.equipo) || '—'),
          td(x.r.tecnico || '—'),
          td(detalleCorto(x.et, x.r))
        ]));
      });
      t.appendChild(tb);
      wrap.appendChild(t);
      card2.appendChild(wrap);
    }
    contentEl.appendChild(card2);
  }

  function detalleCorto(etapa, r) {
    var partes = [];
    if (r.via) partes.push(r.via);
    if (r.estado) partes.push(r.estado);
    if (r.estado_equipo) partes.push(r.estado_equipo);
    if (r.resultado) partes.push(r.resultado);
    if (r.estado_final) partes.push(r.estado_final);
    if (r.tipo_compra) partes.push(r.tipo_compra);
    if (r.empresa) partes.push(r.empresa);
    return partes.join(' · ') || '—';
  }
  function th(t) { return el('th', {}, t); }
  function td(c) { return el('td', {}, [typeof c === 'object' && c ? c : document.createTextNode(c == null ? '' : String(c))]); }

  // --------------------------------------------------------------- Etapa view
  function renderEtapa(id) {
    var etapa = ETAPAS_BY_ID[id];
    setTitulo(etapa.icono + ' ' + etapa.nombre, etapa.via + ' · ' + (DB.registros[id].length) + ' registro(s)');
    contentEl.innerHTML = '';

    var editando = STATE.editId ? DB.registros[id].filter(function (r) { return r._id === STATE.editId; })[0] : null;

    // --- Formulario
    var card = el('div', { class: 'card' });
    card.appendChild(el('div', { class: 'card-head' }, [
      el('h3', {}, (editando ? 'Editar registro' : 'Nuevo registro') + ' · ' + etapa.nombre),
      el('span', { class: 'desc' }, etapa.desc)
    ]));
    var body = el('div', { class: 'card-body' });
    var form = buildForm(etapa, editando);
    body.appendChild(form.grid);

    var actions = el('div', { class: 'form-actions' });
    var btnGuardar = el('button', { class: 'btn btn-primary' }, editando ? '💾 Actualizar registro' : '➕ Guardar registro');
    btnGuardar.onclick = function () {
      try {
        var rec = collectForm(etapa, form.controls);
        if (editando) {
          rec._id = editando._id; rec._stage = id; rec._createdAt = editando._createdAt; rec._updatedAt = new Date().toISOString();
          var idx = DB.registros[id].findIndex(function (r) { return r._id === editando._id; });
          DB.registros[id][idx] = rec;
          toast('Registro actualizado.', 'ok');
        } else {
          rec._id = uid(); rec._stage = id; rec._createdAt = new Date().toISOString();
          DB.registros[id].push(rec);
          toast('Registro de «' + etapa.nombre + '» creado.', 'ok');
        }
        autoaprenderEmpresa(rec);
        guardarDB();
        STATE.editId = null;
        renderEtapa(id);
        window.scrollTo({ top: 0, behavior: 'smooth' });
      } catch (err) {
        toast(err.message, 'err');
      }
    };
    actions.appendChild(btnGuardar);

    if (editando) {
      var btnCancelar = el('button', { class: 'btn' }, 'Cancelar edición');
      btnCancelar.onclick = function () { STATE.editId = null; renderEtapa(id); };
      actions.appendChild(btnCancelar);
    } else {
      var btnLimpiar = el('button', { class: 'btn btn-ghost' }, 'Limpiar');
      btnLimpiar.onclick = function () { renderEtapa(id); };
      actions.appendChild(btnLimpiar);
    }
    body.appendChild(actions);
    card.appendChild(body);
    contentEl.appendChild(card);

    // --- Tabla de registros de la etapa
    var card2 = el('div', { class: 'card' });
    card2.appendChild(el('div', { class: 'card-head' }, [
      el('h3', {}, 'Registros de esta etapa'),
      el('span', { class: 'desc' }, DB.registros[id].length + ' en total')
    ]));
    var body2 = el('div', { class: 'card-body' });

    var toolbar = el('div', { class: 'toolbar' });
    var search = el('input', { type: 'search', placeholder: 'Filtrar registros…' });
    toolbar.appendChild(search);
    toolbar.appendChild(el('div', { class: 'spacer' }));
    var btnExpEt = el('button', { class: 'btn' }, '⬇️ Exportar esta etapa');
    btnExpEt.onclick = function () { exportarEtapa(etapa); };
    toolbar.appendChild(btnExpEt);
    body2.appendChild(toolbar);

    var tablaCont = el('div');
    body2.appendChild(tablaCont);
    function pintarTabla() { tablaCont.innerHTML = ''; tablaCont.appendChild(buildTablaEtapa(etapa, search.value.trim().toLowerCase())); }
    search.addEventListener('input', pintarTabla);
    pintarTabla();

    card2.appendChild(body2);
    contentEl.appendChild(card2);
  }

  function recordHaystack(etapa, r) {
    var s = [r.folio, r.tecnico, fmtFecha(r.fecha), equipoCorto(r.equipo)];
    etapa.campos.forEach(function (c) {
      if (c.tipo === 'equipo' || c.key === 'folio' || c.key === 'tecnico' || c.key === 'fecha') return;
      if (r[c.key]) s.push(r[c.key]);
    });
    return s.join(' ').toLowerCase();
  }

  function buildTablaEtapa(etapa, filtro) {
    var registros = DB.registros[etapa.id].slice();
    registros.sort(function (a, b) {
      return cmpNat(a.folio, b.folio) || cmpNat(a.fecha, b.fecha) || (a._createdAt || '').localeCompare(b._createdAt || '');
    });
    if (filtro) registros = registros.filter(function (r) { return recordHaystack(etapa, r).indexOf(filtro) >= 0; });

    if (!registros.length) {
      return el('div', { class: 'empty-state' }, [el('div', { class: 'big' }, '📭'),
        el('div', {}, filtro ? 'No hay registros que coincidan con el filtro.' : 'Todavía no hay registros en esta etapa.')]);
    }

    // Columnas visibles: folio, fecha, equipo, demás campos (sin observaciones), observaciones, acciones
    var colsCampos = etapa.campos.filter(function (c) { return c.key !== 'observaciones'; });

    var wrap = el('div', { class: 'tabla-wrap' });
    var t = el('table', { class: 'data' });
    var headRow = el('tr');
    colsCampos.forEach(function (c) { headRow.appendChild(th(c.tipo === 'equipo' ? 'Equipo' : c.label)); });
    if (tieneObs(etapa)) headRow.appendChild(th('Observaciones'));
    headRow.appendChild(th('Acciones'));
    t.appendChild(el('thead', {}, headRow));

    var tb = el('tbody');
    registros.forEach(function (r) {
      var tr = el('tr');
      colsCampos.forEach(function (c) {
        if (c.tipo === 'equipo') tr.appendChild(td(equipoCorto(r.equipo) || '—'));
        else if (c.tipo === 'fecha') tr.appendChild(td(fmtFecha(r[c.key]) || '—'));
        else if (['via', 'estado', 'estado_equipo', 'resultado', 'estado_final', 'tipo_compra'].indexOf(c.key) >= 0)
          tr.appendChild(td(r[c.key] ? pillFor(c.key, r[c.key]) : '—'));
        else tr.appendChild(td(r[c.key] || '—'));
      });
      if (tieneObs(etapa)) {
        var obs = r.observaciones || '';
        tr.appendChild(td(obs.length > 60 ? (obs.slice(0, 60) + '…') : (obs || '—')));
      }
      var acc = el('td', { class: 'acciones' });
      var bEd = el('button', { class: 'btn btn-sm' }, '✏️ Editar');
      bEd.onclick = function () { STATE.editId = r._id; renderEtapa(etapa.id); window.scrollTo({ top: 0, behavior: 'smooth' }); };
      var bDel = el('button', { class: 'btn btn-sm btn-danger' }, '🗑️');
      bDel.onclick = function () {
        if (!confirm('¿Eliminar este registro de «' + etapa.nombre + '»? Esta acción no se puede deshacer.')) return;
        DB.registros[etapa.id] = DB.registros[etapa.id].filter(function (x) { return x._id !== r._id; });
        guardarDB(); toast('Registro eliminado.'); renderEtapa(etapa.id);
      };
      acc.appendChild(bEd); acc.appendChild(document.createTextNode(' ')); acc.appendChild(bDel);
      tr.appendChild(acc);
      tb.appendChild(tr);
    });
    t.appendChild(tb);
    wrap.appendChild(t);
    return wrap;
  }
  function tieneObs(etapa) { return etapa.campos.some(function (c) { return c.key === 'observaciones'; }); }

  // ------------------------------------------------------------ Todos los reg.
  function renderTodos() {
    setTitulo('🗂️ Todos los registros', totalRegistros() + ' registros · bitácora general');
    contentEl.innerHTML = '';

    var card = el('div', { class: 'card' });
    var body = el('div', { class: 'card-body' });

    var toolbar = el('div', { class: 'toolbar' });
    var search = el('input', { type: 'search', placeholder: 'Buscar en todos los registros…' });
    var selEtapa = el('select');
    selEtapa.appendChild(el('option', { value: '' }, 'Todas las etapas'));
    ETAPAS.forEach(function (e) { selEtapa.appendChild(el('option', { value: e.id }, e.nombre)); });
    toolbar.appendChild(search);
    toolbar.appendChild(selEtapa);
    toolbar.appendChild(el('div', { class: 'spacer' }));
    var note = el('span', { class: 'count-note' });
    toolbar.appendChild(note);
    body.appendChild(toolbar);

    var cont = el('div');
    body.appendChild(cont);

    function pintar() {
      var rows = bitacoraRows();
      var f = search.value.trim().toLowerCase();
      var etf = selEtapa.value;
      rows = rows.filter(function (x) {
        if (etf && x.et.id !== etf) return false;
        if (!f) return true;
        return x._hay.indexOf(f) >= 0;
      });
      note.textContent = rows.length + ' registro(s)';
      cont.innerHTML = '';
      if (!rows.length) { cont.appendChild(el('div', { class: 'empty-state' }, [el('div', { class: 'big' }, '🔎'), el('div', {}, 'Sin resultados.')])); return; }

      var wrap = el('div', { class: 'tabla-wrap' });
      var t = el('table', { class: 'data' });
      t.appendChild(el('thead', {}, el('tr', {}, [
        th('Folio'), th('Fecha'), th('Etapa'), th('Vía'), th('Inventario'), th('Equipo'), th('Servicio'),
        th('Técnico'), th('Estado / Resultado'), th('Empresa'), th('N° documento'), th('Acciones')
      ])));
      var tb = el('tbody');
      rows.forEach(function (x) {
        var r = x.r;
        var tr = el('tr');
        tr.appendChild(td(r.folio || '—'));
        tr.appendChild(td(fmtFecha(r.fecha) || '—'));
        tr.appendChild(td(el('span', { class: 'tag-etapa' }, x.et.nombre)));
        tr.appendChild(td(x.et.via));
        tr.appendChild(td(r.equipo ? (r.equipo.inv || '—') : '—'));
        tr.appendChild(td(r.equipo ? (r.equipo.nombre || '—') : '—'));
        tr.appendChild(td(r.equipo ? (r.equipo.servicio || '—') : '—'));
        tr.appendChild(td(r.tecnico || '—'));
        tr.appendChild(td(estadoResultado(r) ? pillFor(estadoResultadoKey(r), estadoResultado(r)) : '—'));
        tr.appendChild(td(r.empresa || '—'));
        tr.appendChild(td(numeroDoc(r) || '—'));
        var acc = el('td', { class: 'acciones' });
        var bEd = el('button', { class: 'btn btn-sm' }, '✏️');
        bEd.onclick = function () { STATE.editId = r._id; navegar(x.et.id); window.scrollTo({ top: 0, behavior: 'smooth' }); };
        acc.appendChild(bEd);
        tr.appendChild(acc);
        tb.appendChild(tr);
      });
      t.appendChild(tb);
      wrap.appendChild(t);
      cont.appendChild(wrap);
    }
    search.addEventListener('input', pintar);
    selEtapa.addEventListener('change', pintar);
    pintar();

    card.appendChild(body);
    contentEl.appendChild(card);
  }

  function bitacoraRows() {
    var rows = [];
    ETAPAS.forEach(function (et) {
      DB.registros[et.id].forEach(function (r) {
        rows.push({
          et: et, r: r,
          _hay: [r.folio, fmtFecha(r.fecha), et.nombre, et.via, r.equipo && r.equipo.inv, r.equipo && r.equipo.nombre,
            r.equipo && r.equipo.servicio, r.tecnico, estadoResultado(r), r.empresa, numeroDoc(r), r.observaciones]
            .filter(Boolean).join(' ').toLowerCase()
        });
      });
    });
    rows.sort(function (a, b) {
      return cmpNat(a.r.folio, b.r.folio) || cmpNat(a.r.fecha, b.r.fecha) || (a.et._orden - b.et._orden) ||
        (a.r._createdAt || '').localeCompare(b.r._createdAt || '');
    });
    return rows;
  }
  function estadoResultado(r) { return r.estado || r.estado_equipo || r.resultado || r.estado_final || r.via || ''; }
  function estadoResultadoKey(r) {
    if (r.estado) return 'estado'; if (r.estado_equipo) return 'estado_equipo'; if (r.resultado) return 'resultado';
    if (r.estado_final) return 'estado_final'; if (r.via) return 'via'; return '';
  }
  function numeroDoc(r) { return r.numero_envio || r.numero_guia || r.numero_cotizacion || r.numero_informe || r.numero_oc || ''; }

  // ----------------------------------------------------------------- Config
  function renderConfig() {
    setTitulo('⚙️ Configuración', 'Técnicos, empresas y datos');
    contentEl.innerHTML = '';

    // Técnicos
    var card = el('div', { class: 'card' });
    card.appendChild(el('div', { class: 'card-head' }, [el('h3', {}, '👷 Técnicos'), el('span', { class: 'desc' }, 'Aparecen en la lista desplegable de todas las etapas.')]));
    var body = el('div', { class: 'card-body' });
    body.appendChild(el('div', { class: 'hint' }, 'Supervisión: ' + SUPERVISOR + '.'));
    var chips = el('div', { class: 'chips' });
    DB.config.tecnicos.forEach(function (t) {
      var chip = el('div', { class: 'chip' }, [el('span', {}, t), el('span', { class: 'x', title: 'Quitar' }, '✕')]);
      chip.querySelector('.x').onclick = function () {
        DB.config.tecnicos = DB.config.tecnicos.filter(function (x) { return x !== t; });
        guardarDB(); renderConfig();
      };
      chips.appendChild(chip);
    });
    body.appendChild(chips);
    var add = el('div', { class: 'inline-add' });
    var inp = el('input', { type: 'text', placeholder: 'Nombre del técnico…' });
    var btn = el('button', { class: 'btn btn-primary' }, 'Agregar');
    function addTec() {
      var v = inp.value.trim(); if (!v) return;
      if (DB.config.tecnicos.indexOf(v) < 0) DB.config.tecnicos.push(v);
      DB.config.tecnicos.sort(cmpNat); guardarDB(); renderConfig();
    }
    btn.onclick = addTec; inp.addEventListener('keydown', function (e) { if (e.key === 'Enter') addTec(); });
    add.appendChild(inp); add.appendChild(btn);
    body.appendChild(add);
    card.appendChild(body);
    contentEl.appendChild(card);

    // Empresas
    var card2 = el('div', { class: 'card' });
    card2.appendChild(el('div', { class: 'card-head' }, [el('h3', {}, '🏢 Empresas / proveedores'), el('span', { class: 'desc' }, 'Sugerencias para los campos de empresa.')]));
    var body2 = el('div', { class: 'card-body' });
    var chips2 = el('div', { class: 'chips' });
    if (!DB.config.empresas.length) chips2.appendChild(el('div', { class: 'hint' }, 'Aún no hay empresas. Se agregan automáticamente al registrarlas o puede añadirlas aquí.'));
    DB.config.empresas.slice().sort(cmpNat).forEach(function (t) {
      var chip = el('div', { class: 'chip' }, [el('span', {}, t), el('span', { class: 'x', title: 'Quitar' }, '✕')]);
      chip.querySelector('.x').onclick = function () {
        DB.config.empresas = DB.config.empresas.filter(function (x) { return x !== t; });
        guardarDB(); renderConfig();
      };
      chips2.appendChild(chip);
    });
    body2.appendChild(chips2);
    var add2 = el('div', { class: 'inline-add' });
    var inp2 = el('input', { type: 'text', placeholder: 'Nombre de la empresa…' });
    var btn2 = el('button', { class: 'btn btn-primary' }, 'Agregar');
    function addEmp() {
      var v = inp2.value.trim(); if (!v) return;
      if (DB.config.empresas.indexOf(v) < 0) DB.config.empresas.push(v);
      guardarDB(); renderConfig();
    }
    btn2.onclick = addEmp; inp2.addEventListener('keydown', function (e) { if (e.key === 'Enter') addEmp(); });
    add2.appendChild(inp2); add2.appendChild(btn2);
    body2.appendChild(add2);
    card2.appendChild(body2);
    contentEl.appendChild(card2);

    // Datos
    var card3 = el('div', { class: 'card' });
    card3.appendChild(el('div', { class: 'card-head' }, [el('h3', {}, '💾 Datos y respaldo'), el('span', { class: 'desc' }, EQUIPOS.length + ' equipos críticos cargados como referencia.')]));
    var body3 = el('div', { class: 'card-body' });
    body3.appendChild(el('div', { class: 'hint' }, 'Los registros se guardan localmente en este navegador (localStorage). Use el respaldo para trasladarlos a otro equipo.'));
    var actions = el('div', { class: 'form-actions' });

    var bExport = el('button', { class: 'btn btn-success' }, '⬇️ Exportar a Excel (.xlsx)');
    bExport.onclick = exportarTodo;
    var bBackup = el('button', { class: 'btn' }, '🗄️ Descargar respaldo (JSON)');
    bBackup.onclick = exportarRespaldo;
    var bRestore = el('button', { class: 'btn' }, '📤 Restaurar respaldo (JSON)');
    bRestore.onclick = importarRespaldo;
    var bClear = el('button', { class: 'btn btn-danger' }, '🗑️ Borrar todos los registros');
    bClear.onclick = function () {
      if (!confirm('¿Borrar TODOS los registros? Esta acción no se puede deshacer. Se recomienda descargar un respaldo antes.')) return;
      ETAPAS.forEach(function (e) { DB.registros[e.id] = []; });
      guardarDB(); toast('Registros eliminados.'); navegar('__dashboard');
    };
    actions.appendChild(bExport); actions.appendChild(bBackup); actions.appendChild(bRestore); actions.appendChild(bClear);
    body3.appendChild(actions);
    card3.appendChild(body3);
    contentEl.appendChild(card3);
  }

  function autoaprenderEmpresa(rec) {
    if (rec.empresa && DB.config.empresas.indexOf(rec.empresa) < 0) DB.config.empresas.push(rec.empresa);
  }

  // ------------------------------------------------------------- Respaldo JSON
  function exportarRespaldo() {
    var blob = new Blob([JSON.stringify(DB, null, 2)], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: 'Respaldo_Gestion_Equipos_' + hoyISO() + '.json' });
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
    toast('Respaldo descargado.', 'ok');
  }
  function importarRespaldo() {
    var inp = el('input', { type: 'file', accept: '.json,application/json' });
    inp.onchange = function () {
      var file = inp.files[0]; if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var data = JSON.parse(reader.result);
          if (!data.registros) throw new Error('Archivo no válido.');
          if (!confirm('Esto reemplazará los registros actuales por los del respaldo. ¿Continuar?')) return;
          DB = data;
          DB.config = DB.config || {};
          if (!DB.config.tecnicos || !DB.config.tecnicos.length) DB.config.tecnicos = TECNICOS_DEFAULT.slice();
          if (!DB.config.empresas) DB.config.empresas = [];
          ETAPAS.forEach(function (e) { if (!Array.isArray(DB.registros[e.id])) DB.registros[e.id] = []; });
          guardarDB(); toast('Respaldo restaurado.', 'ok'); navegar('__dashboard');
        } catch (e) { toast('No se pudo leer el respaldo: ' + e.message, 'err'); }
      };
      reader.readAsText(file);
    };
    inp.click();
  }

  // ----------------------------------------------------------- Export a Excel
  function columnasExport(etapa) {
    var cols = [];
    etapa.campos.forEach(function (c) {
      if (c.tipo === 'equipo') {
        cols.push({ titulo: 'N° Inventario', ancho: 14, get: function (r) { return r[c.key] ? (r[c.key].inv || '') : ''; } });
        cols.push({ titulo: 'Equipo', ancho: 20, get: function (r) { return r[c.key] ? (r[c.key].nombre || '') : ''; } });
        cols.push({ titulo: 'Servicio', ancho: 22, get: function (r) { return r[c.key] ? (r[c.key].servicio || '') : ''; } });
        cols.push({ titulo: 'Unidad', ancho: 18, get: function (r) { return r[c.key] ? (r[c.key].unidad || '') : ''; } });
        cols.push({ titulo: 'Ubicación', ancho: 18, get: function (r) { return r[c.key] ? (r[c.key].ubicacion || '') : ''; } });
        cols.push({ titulo: 'Marca / Modelo', ancho: 22, get: function (r) { return r[c.key] ? [r[c.key].marca, r[c.key].modelo].filter(Boolean).join(' ') : ''; } });
        cols.push({ titulo: 'Serie', ancho: 16, get: function (r) { return r[c.key] ? (r[c.key].serie || '') : ''; } });
      } else {
        var label = c.label;
        var key = c.key;
        cols.push({
          titulo: label, ancho: c.ancho || 16, get: (function (kk, tipo) {
            return function (r) { return tipo === 'fecha' ? fmtFecha(r[kk]) : (r[kk] != null ? r[kk] : ''); };
          })(key, c.tipo)
        });
      }
    });
    cols.push({ titulo: 'Registrado el', ancho: 18, get: function (r) { return fmtFechaHora(r._createdAt); } });
    return cols;
  }

  function hojaEtapa(etapa) {
    var registros = DB.registros[etapa.id].slice();
    registros.sort(function (a, b) {
      return cmpNat(a.folio, b.folio) || cmpNat(a.fecha, b.fecha) || (a._createdAt || '').localeCompare(b._createdAt || '');
    });
    var cols = columnasExport(etapa);
    return {
      nombre: etapa.nombre,
      columnas: cols.map(function (c) { return { titulo: c.titulo, ancho: c.ancho }; }),
      filas: registros.map(function (r) { return cols.map(function (c) { return c.get(r); }); })
    };
  }

  function hojaBitacora() {
    var rows = bitacoraRows();
    var cols = [
      { titulo: 'Folio', ancho: 14, get: function (x) { return x.r.folio || ''; } },
      { titulo: 'Fecha', ancho: 12, get: function (x) { return fmtFecha(x.r.fecha); } },
      { titulo: 'Etapa', ancho: 26, get: function (x) { return x.et.nombre; } },
      { titulo: 'Vía', ancho: 12, get: function (x) { return x.et.via; } },
      { titulo: 'N° Inventario', ancho: 14, get: function (x) { return x.r.equipo ? (x.r.equipo.inv || '') : ''; } },
      { titulo: 'Equipo', ancho: 20, get: function (x) { return x.r.equipo ? (x.r.equipo.nombre || '') : ''; } },
      { titulo: 'Servicio', ancho: 22, get: function (x) { return x.r.equipo ? (x.r.equipo.servicio || '') : ''; } },
      { titulo: 'Serie', ancho: 16, get: function (x) { return x.r.equipo ? (x.r.equipo.serie || '') : ''; } },
      { titulo: 'Técnico', ancho: 24, get: function (x) { return x.r.tecnico || ''; } },
      { titulo: 'Estado / Resultado', ancho: 18, get: function (x) { return estadoResultado(x.r); } },
      { titulo: 'Empresa', ancho: 22, get: function (x) { return x.r.empresa || ''; } },
      { titulo: 'N° documento', ancho: 16, get: function (x) { return numeroDoc(x.r); } },
      { titulo: 'Observaciones', ancho: 40, get: function (x) { return x.r.observaciones || ''; } },
      { titulo: 'Registrado el', ancho: 18, get: function (x) { return fmtFechaHora(x.r._createdAt); } }
    ];
    return {
      nombre: 'Bitácora general',
      columnas: cols.map(function (c) { return { titulo: c.titulo, ancho: c.ancho }; }),
      filas: rows.map(function (x) { return cols.map(function (c) { return c.get(x); }); })
    };
  }

  function exportarTodo() {
    if (totalRegistros() === 0) { toast('No hay registros para exportar.', 'err'); return; }
    var hojas = [hojaBitacora()];
    ETAPAS.forEach(function (et) { if (DB.registros[et.id].length) hojas.push(hojaEtapa(et)); });
    try {
      XLSXWriter.descargar('Registros_Gestion_Equipos_' + hoyISO() + '.xlsx', hojas);
      toast('Exportado a Excel (' + hojas.length + ' hojas).', 'ok');
    } catch (e) { toast('Error al exportar: ' + e.message, 'err'); }
  }

  function exportarEtapa(etapa) {
    if (!DB.registros[etapa.id].length) { toast('No hay registros en esta etapa.', 'err'); return; }
    try {
      XLSXWriter.descargar(etapa.nombre.replace(/[^\wáéíóúñ ]/gi, '').trim().replace(/\s+/g, '_') + '_' + hoyISO() + '.xlsx', [hojaEtapa(etapa)]);
      toast('Etapa exportada a Excel.', 'ok');
    } catch (e) { toast('Error al exportar: ' + e.message, 'err'); }
  }

  // ------------------------------------------------------------------- Init
  function init() {
    contentEl = document.getElementById('content');
    viewTitleEl = document.getElementById('viewTitle');
    viewSubEl = document.getElementById('viewSub');

    document.getElementById('btnExport').onclick = exportarTodo;
    var mt = document.getElementById('menuToggle');
    var sb = document.getElementById('sidebar');
    var bd = document.getElementById('backdrop');
    mt.onclick = function () { sb.classList.toggle('open'); bd.classList.toggle('show'); };
    bd.onclick = function () { sb.classList.remove('open'); bd.classList.remove('show'); };

    render();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
