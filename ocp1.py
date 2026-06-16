<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard KPI — Performance & Qualité</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" rel="stylesheet">
<script>
tailwind.config={theme:{extend:{fontFamily:{main:['Space Grotesk','sans-serif']}}}}
</script>
<style>
:root{
  --hdr:#0f2d3d;--hdr2:#164e63;--fcol:#e0f2fe;--fcol-t:#0c4a6e;
  --tot-bg:#d1fae5;--tot-t:#065f46;--page:#f0f4f8;
  --perf:#059669;--qual:#0284c7;--ano:#dc2626;--bl:#7c3aed;
  --good:#c6efce;--good-t:#006100;--warn:#ffeb9c;--warn-t:#9c6500;
  --bad:#ffc7ce;--bad-t:#9c0006;
  --sc-bg:#f0fdf4;--sc-bdr:#059669;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Space Grotesk',sans-serif;background:var(--page);min-height:100vh;display:flex}

/* ===== SIDEBAR ===== */
.sidebar{width:244px;min-height:100vh;background:linear-gradient(180deg,#071a28 0%,#0f2d3d 35%,#134e5e 100%);position:fixed;left:0;top:0;z-index:50;display:flex;flex-direction:column;box-shadow:4px 0 24px rgba(0,0,0,.18)}
.sidebar-brand{padding:20px 16px 16px;border-bottom:1px solid rgba(255,255,255,.07)}
.sidebar-brand h2{color:#fff;font-size:15px;font-weight:700;letter-spacing:.3px;display:flex;align-items:center;gap:8px}
.sidebar-brand h2 i{color:#14b8a6;font-size:16px}
.sidebar-brand p{color:rgba(255,255,255,.38);font-size:10.5px;margin-top:3px;letter-spacing:.5px}
.sidebar-nav{flex:1;padding:12px 8px;overflow-y:auto}
.nav-section-label{color:rgba(255,255,255,.3);font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;padding:14px 14px 5px}
.nav-item{display:flex;align-items:center;gap:10px;padding:9px 14px;border-radius:8px;color:rgba(255,255,255,.65);font-size:13px;font-weight:500;cursor:pointer;transition:all .2s;margin-bottom:1px;text-decoration:none;position:relative;user-select:none}
.nav-item:hover{background:rgba(255,255,255,.07);color:rgba(255,255,255,.95)}
.nav-item.active{background:rgba(13,148,136,.22);color:#5eead4;font-weight:600}
.nav-item.active::before{content:'';position:absolute;left:0;top:50%;transform:translateY(-50%);width:3px;height:20px;background:#14b8a6;border-radius:0 3px 3px 0}
.nav-item i{width:18px;text-align:center;font-size:13px}
.sidebar-footer{padding:14px 16px;border-top:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.28);font-size:10px;line-height:1.6}
.sidebar-footer .hse-badge{display:inline-flex;align-items:center;gap:5px;color:#14b8a6;font-weight:600;margin-bottom:2px}

/* ===== MAIN ===== */
.main{margin-left:244px;flex:1;padding:16px 22px;min-height:100vh;position:relative}
.main::before{content:'';position:fixed;top:0;left:244px;right:0;bottom:0;background:radial-gradient(ellipse at 95% 5%,rgba(13,148,136,.035),transparent 55%),radial-gradient(ellipse at 5% 95%,rgba(5,150,105,.025),transparent 55%);pointer-events:none;z-index:0}
.main>*{position:relative;z-index:1}

/* ===== PAGE HEADER ===== */
.page-header{background:linear-gradient(135deg,var(--hdr),var(--hdr2));padding:14px 22px;border-radius:12px;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 6px 24px rgba(0,0,0,.12);position:relative;overflow:hidden}
.page-header::after{content:'';position:absolute;top:-40px;right:-40px;width:140px;height:140px;background:rgba(255,255,255,.025);border-radius:50%}
.page-header::before{content:'';position:absolute;bottom:-20px;left:30%;width:80px;height:80px;background:rgba(255,255,255,.015);border-radius:50%}
.page-header h1{color:#fff;font-size:20px;font-weight:800;position:relative;display:flex;align-items:center;gap:8px}
.page-header h1 i{font-size:18px;opacity:.8}
.page-header .badge{background:rgba(255,255,255,.11);padding:5px 16px;border-radius:18px;color:rgba(255,255,255,.9);font-size:13px;font-weight:500;border:1px solid rgba(255,255,255,.12);display:flex;align-items:center;gap:5px;position:relative}

/* ===== CARDS ===== */
.cards-row{display:grid;gap:10px;margin-bottom:14px}
.cards-4{grid-template-columns:repeat(4,1fr)}
.cards-3{grid-template-columns:repeat(3,1fr)}
.cards-2{grid-template-columns:repeat(2,1fr)}
.kpi-card{background:#fff;border-radius:10px;padding:14px 14px 12px;box-shadow:0 2px 8px rgba(0,0,0,.04);border:1px solid #e2e8f0;text-align:center;transition:transform .2s,box-shadow .2s;position:relative;overflow:hidden}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(0,0,0,.07)}
.kpi-card .val{font-size:28px;font-weight:900;line-height:1.1;display:flex;align-items:baseline;justify-content:center;gap:2px}
.kpi-card .val .icon{font-size:16px;opacity:.5;font-weight:400}
.kpi-card .lbl{font-size:10.5px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-top:4px;line-height:1.3}
.kpi-card.perf{border-top:3px solid var(--perf)}.kpi-card.perf .val{color:var(--perf)}
.kpi-card.qual{border-top:3px solid var(--qual)}.kpi-card.qual .val{color:var(--qual)}
.kpi-card.ano{border-top:3px solid var(--ano)}.kpi-card.ano .val{color:var(--ano)}
.kpi-card.bl{border-top:3px solid var(--bl)}.kpi-card.bl .val{color:var(--bl)}
.kpi-card.neutral{border-top:3px solid #475569}.kpi-card.neutral .val{color:#334155}

/* ===== GROUP LABEL ===== */
.group-label{display:inline-flex;align-items:center;gap:8px;padding:7px 18px;border-radius:8px;font-size:13px;font-weight:700;margin:14px 0 8px 0;letter-spacing:.3px;color:#fff;background:linear-gradient(135deg,var(--hdr),var(--hdr2));box-shadow:0 2px 8px rgba(0,0,0,.08)}
.group-label .dot{width:8px;height:8px;border-radius:50%;background:#5eead4;animation:pulse-dot 2s ease infinite}
@keyframes pulse-dot{0%,100%{opacity:1}50%{opacity:.4}}
.group-label .posts-list{font-weight:400;opacity:.55;font-size:12px}

/* ===== SECTION TITLE ===== */
.section-title{font-size:15px;font-weight:700;color:var(--hdr);margin:18px 0 8px 0;padding-left:10px;border-left:3px solid var(--hdr2);display:flex;align-items:center;gap:7px}
.section-title.perf-t{border-left-color:var(--perf)}
.section-title.qual-t{border-left-color:var(--qual)}
.section-title.ano-t{border-left-color:var(--ano)}
.section-title.bl-t{border-left-color:var(--bl)}
.section-title i{font-size:14px;opacity:.7}

/* ===== TABLES ===== */
.tbl-wrap{overflow-x:auto;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.04);margin-bottom:12px}
.kpi-tbl{width:100%;border-collapse:collapse;font-size:12px}
.kpi-tbl thead th{background:linear-gradient(135deg,var(--hdr),var(--hdr2));color:#fff;font-weight:700;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;padding:8px 8px;border:none;white-space:nowrap;position:sticky;top:0;z-index:5}
.kpi-tbl thead th.fc-head{background:linear-gradient(135deg,#0c3a4f,#1a5c72);text-align:left;min-width:140px}
.kpi-tbl thead th.sc-head{background:linear-gradient(135deg,#134e4a,#0f766e);border-left:2px solid rgba(255,255,255,.15)}
.kpi-tbl thead th.act-head{background:linear-gradient(135deg,#991b1b,#dc2626)}
.kpi-tbl thead th.ano-head{background:linear-gradient(135deg,#991b1b,#dc2626)}
.kpi-tbl thead th.bl-head{background:linear-gradient(135deg,#4c1d95,#7c3aed)}
.kpi-tbl tbody td{padding:5px 8px;border-bottom:1px solid #f1f5f9;white-space:nowrap;text-align:center}
.kpi-tbl tbody td.fc{background:var(--fcol)!important;color:var(--fcol-t);font-weight:700;text-align:left;min-width:140px;position:sticky;left:0;z-index:2;box-shadow:2px 0 4px rgba(0,0,0,.03)}
.kpi-tbl tbody td.sc{background:var(--sc-bg)!important;border-left:2px solid var(--sc-bdr);font-weight:800}
.kpi-tbl tbody tr:nth-child(even) td:not(.fc):not(.sc){background:#f8fafc}
.kpi-tbl tbody tr:hover td:not(.fc):not(.sc){background:#f0fdfa!important}
.kpi-tbl tbody tr.total-row td{background:var(--tot-bg)!important;color:var(--tot-t)!important;font-weight:800!important;font-size:12.5px!important}
.kpi-tbl tbody tr.total-row td.fc{background:#a7f3d0!important;color:#064e3b!important;box-shadow:2px 0 4px rgba(0,0,0,.05)}
.kpi-tbl tbody tr.total-row td.sc{background:#6ee7b7!important;border-left-color:#047857;color:#064e3b!important}

/* ===== CELL CLASSES ===== */
.cg{background:var(--good)!important;color:var(--good-t)!important;font-weight:600}
.cw{background:var(--warn)!important;color:var(--warn-t)!important;font-weight:600}
.cb{background:var(--bad)!important;color:var(--bad-t)!important;font-weight:600}
.a0{background:var(--good)!important;color:var(--good-t)!important;font-weight:600}
.a1{background:var(--warn)!important;color:var(--warn-t)!important;font-weight:600}
.a2{background:#fed7d7!important;color:#c53030!important;font-weight:600}
.a3{background:#fecaca!important;color:#991b1b!important;font-weight:800}

/* ===== CHART BOXES ===== */
.chart-box{background:#fff;border-radius:10px;padding:14px;border:1px solid #e2e8f0;box-shadow:0 2px 8px rgba(0,0,0,.04)}
.chart-box .chart-title{font-size:13px;font-weight:700;margin-bottom:8px;color:var(--hdr);padding-bottom:6px;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;gap:6px}
.chart-box .chart-title i{font-size:12px;opacity:.5}

/* ===== PROGRESS BAR ===== */
.pbar{height:6px;background:#e2e8f0;border-radius:3px;overflow:hidden}
.pbar-fill{height:100%;border-radius:3px;transition:width .6s ease}

/* ===== MINI BAR GROUP ===== */
.mini-bars{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.mini-bar-item{display:flex;align-items:center;gap:8px;font-size:12px;padding:5px 8px;background:#f8fafc;border-radius:6px;border:1px solid #f1f5f9}
.mini-bar-item .poste-name{min-width:38px;font-weight:700;color:var(--fcol-t)}
.mini-bar-item .bars-col{flex:1;display:flex;flex-direction:column;gap:3px}
.mini-bar-row{display:flex;align-items:center;gap:4px}
.mini-bar-row .bar-label{font-size:9px;color:#94a3b8;min-width:18px;font-weight:600;text-transform:uppercase}
.mini-bar-row .pbar{flex:1}
.mini-bar-row .bar-val{min-width:42px;text-align:right;font-weight:700;font-size:11px}

/* ===== PAGE SECTIONS ===== */
.page-section{display:none}
.page-section.active{display:block}

/* ===== EVOLUTION TIMELINE (Suivi masqué) ===== */
.evol-timeline{border-left:3px solid #0f2d3d;margin-left:12px;padding-left:22px}
.evol-item{position:relative;padding-bottom:22px}
.evol-item::before{content:'';position:absolute;left:-28px;top:4px;width:12px;height:12px;border-radius:50%;background:#0d9488;border:2px solid #fff;box-shadow:0 0 0 2px #0d9488}
.evol-ver{font-size:14px;font-weight:800;color:#0f2d3d}
.evol-date{font-size:11px;color:#64748b;margin:2px 0 6px}
.evol-change{font-size:12px;color:#475569;padding:2px 0 2px 14px;position:relative;line-height:1.5}
.evol-change::before{content:'\2022';position:absolute;left:0;color:#059669;font-weight:800;font-size:14px}

/* ===== SCROLLBAR ===== */
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:#f1f5f9}
::-webkit-scrollbar-thumb{background:#94a3b8;border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:#64748b}

/* ===== ANIMATIONS ===== */
@media(prefers-reduced-motion:no-preference){
  .fade-in{animation:fadeUp .45s ease both}
  .fd1{animation:fadeUp .45s ease .08s both}
  .fd2{animation:fadeUp .45s ease .16s both}
  .fd3{animation:fadeUp .45s ease .24s both}
  .fd4{animation:fadeUp .45s ease .32s both}
}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}

/* ===== RESPONSIVE ===== */
@media(max-width:1280px){.cards-4{grid-template-columns:repeat(2,1fr)}.cards-3{grid-template-columns:repeat(2,1fr)}}
@media(max-width:768px){.sidebar{width:0;overflow:hidden;transition:width .3s}.sidebar.open{width:244px}.main{margin-left:0}.cards-4,.cards-3,.cards-2{grid-template-columns:1fr}.mini-bars{grid-template-columns:1fr}}
</style>
</head>
<body>

<!-- ===== SIDEBAR ===== -->
<aside class="sidebar" id="sidebar">
  <div class="sidebar-brand">
    <h2><i class="fas fa-chart-line"></i>Dashboard KPI</h2>
    <p>Maintenance Industrielle</p>
  </div>
  <nav class="sidebar-nav">
    <div class="nav-section-label">Analyse</div>
    <a class="nav-item active" data-page="dashboard" onclick="navigateTo('dashboard')">
      <i class="fas fa-tachometer-alt"></i>Performance & Qualité
    </a>
    <a class="nav-item" data-page="anomalies" onclick="navigateTo('anomalies')">
      <i class="fas fa-exclamation-triangle"></i>Anomalies
    </a>
    <div class="nav-section-label">Suivi</div>
    <a class="nav-item" data-page="backlog" onclick="navigateTo('backlog')">
      <i class="fas fa-layer-group"></i>Backlog
    </a>
    <!-- ===== PAGE SUIVI MASQUÉE DU MENU — CODE CONSERVÉ POUR RÉACTIVATION =====
    <a class="nav-item" data-page="suivi" onclick="navigateTo('suivi')">
      <i class="fas fa-tasks"></i>Suivi des Améliorations
    </a>
    ===== FIN SUIVI MASQUÉ ===== -->
  </nav>
  <div class="sidebar-footer">
    <div class="hse-badge"><i class="fas fa-shield-alt"></i> HSE — Sécurité active</div>
    <div>Version 2.1 — 18/06/2025</div>
  </div>
</aside>

<!-- ===== MAIN ===== -->
<div class="main">
  <div id="page-dashboard" class="page-section active"></div>
  <div id="page-anomalies" class="page-section"></div>
  <div id="page-backlog" class="page-section"></div>
  <!-- Page Suivi masquée — code conservé intact pour réactivation ultérieure -->
  <div id="page-suivi" class="page-section"></div>
</div>

<script>
// =====================================================================
// CONSTANTES KPI
// =====================================================================
const PERF_KPI = [
  'TAUX_REALISATION_CORRECTIF/PT','OT préparation <1 mois','OT préparation >3 mois',
  'OT préparation 1mois< <3mois','OT planification <1 mois','OT planification >3 mois',
  'OT planification 1mois< <3mois','OT exécution <1 mois','OT exécution >3 mois',
  'OT exécution 1mois< <3mois','Performance Graissage','Performance Inspection',
  'Performance Appels Systématiques'
];
const QUAL_KPI = [
  'appel avis approuvé','OT LANC ESTIME','Backlog préparation caractérisé',
  'Backlog planification caractérisé','OT CONFIME','OT_COR_EGAL',
  'OT Fiabilité','Total Avis de Panne'
];
const ALL_KPI = [...PERF_KPI, ...QUAL_KPI];

const CIBLES = {
  'TAUX_REALISATION_CORRECTIF/PT':85,'OT préparation <1 mois':80,'OT préparation >3 mois':5,
  'OT préparation 1mois< <3mois':15,'OT planification <1 mois':80,'OT planification >3 mois':5,
  'OT planification 1mois< <3mois':15,'OT exécution <1 mois':80,'OT exécution >3 mois':5,
  'OT exécution 1mois< <3mois':15,'appel avis approuvé':95,'OT LANC ESTIME':100,
  'Backlog préparation caractérisé':100,'Backlog planification caractérisé':100,
  'OT CONFIME':100,'OT_COR_EGAL':100,'Performance Graissage':95,'Performance Inspection':95,
  'Performance Appels Systématiques':95,'OT Fiabilité':100,'Total Avis de Panne':100
};

const LOWER_BETTER = [
  'OT préparation >3 mois','OT planification >3 mois','OT exécution >3 mois',
  'OT préparation 1mois< <3mois','OT planification 1mois< <3mois','OT exécution 1mois< <3mois'
];

// Groupes de postes de travail
const GROUPS = {
  sf1: ['SF1'],
  sf2: ['SF2'],
  autres: ['UTA','UTB','ATC','ATD']
};
const ALL_POSTES = [...GROUPS.sf1, ...GROUPS.sf2, ...GROUPS.autres];
const DATE_STR = '18/06/2025';

// =====================================================================
// SEUILS D'ANOMALIE — même logique que gscore() dans le code Python
// =====================================================================
function gscoreMet(k, v) {
  if (LOWER_BETTER.includes(k)) {
    if (k.includes('>3 mois')) return v <= 5;
    if (k.includes('1mois<')) return v <= 15;
  }
  if (k === 'TAUX_REALISATION_CORRECTIF/PT') return v >= 80;
  if (k === 'appel avis approuvé') return v >= 90;
  if (['OT LANC ESTIME','Backlog préparation caractérisé','Backlog planification caractérisé','OT CONFIME','OT_COR_EGAL'].includes(k)) return v >= 95;
  if (['Performance Graissage','Performance Inspection','Performance Appels Systématiques'].includes(k)) return v >= 95;
  if (['OT Fiabilité','Total Avis de Panne'].includes(k)) return v >= 100;
  if (k.includes('<1 mois')) return v >= 75;
  return v >= 80;
}

// =====================================================================
// COULEURS CELLULES — même logique que ks() dans le code Python
// =====================================================================
function getCellClass(k, v) {
  const val = parseFloat(v);
  if (isNaN(val)) return '';
  if (k.includes('OT préparation <1 mois') || k.includes('OT planification <1 mois') || k.includes('OT exécution <1 mois'))
    return val >= 80 ? 'cg' : (val >= 75 ? 'cw' : 'cb');
  if (k.includes('1mois< <3mois'))
    return val <= 15 ? 'cg' : 'cb';
  if (k.includes('>3 mois'))
    return val <= 5 ? 'cg' : 'cb';
  if (k === 'TAUX_REALISATION_CORRECTIF/PT')
    return val >= 85 ? 'cg' : (val >= 80 ? 'cw' : 'cb');
  if (k === 'appel avis approuvé')
    return val >= 95 ? 'cg' : (val >= 90 ? 'cw' : 'cb');
  if (['OT LANC ESTIME','Backlog préparation caractérisé','Backlog planification caractérisé','OT CONFIME','OT_COR_EGAL'].includes(k))
    return val >= 100 ? 'cg' : (val >= 95 ? 'cw' : 'cb');
  if (['Performance Graissage','Performance Inspection','Performance Appels Systématiques'].includes(k))
    return val >= 95 ? 'cg' : (val > 90 ? 'cw' : 'cb');
  if (['OT Fiabilité','Total Avis de Panne'].includes(k))
    return val >= 100 ? 'cg' : (val >= 95 ? 'cw' : 'cb');
  return val >= 80 ? 'cg' : (val >= 70 ? 'cw' : 'cb');
}

function getScoreClass(s) { return s >= 90 ? 'cg' : (s >= 80 ? 'cw' : 'cb'); }
function getAnoCellClass(n) { return n === 0 ? 'a0' : (n <= 1 ? 'a1' : (n <= 3 ? 'a2' : 'a3')); }
function fmt(v) { return parseFloat(v).toFixed(1); }

// =====================================================================
// DONNÉES MOCK — tous les postes intégrés, valeurs réalistes
// =====================================================================
const DATA = {
  'SF1': {
    'TAUX_REALISATION_CORRECTIF/PT':88.5,'OT préparation <1 mois':82.3,'OT préparation >3 mois':3.2,
    'OT préparation 1mois< <3mois':14.5,'OT planification <1 mois':85.1,'OT planification >3 mois':2.8,
    'OT planification 1mois< <3mois':12.1,'OT exécution <1 mois':78.4,'OT exécution >3 mois':6.5,
    'OT exécution 1mois< <3mois':15.1,'Performance Graissage':96.2,'Performance Inspection':94.8,
    'Performance Appels Systématiques':97.1,'appel avis approuvé':96.5,'OT LANC ESTIME':98.2,
    'Backlog préparation caractérisé':97.5,'Backlog planification caractérisé':96.8,
    'OT CONFIME':99.1,'OT_COR_EGAL':95.5,'OT Fiabilité':100,'Total Avis de Panne':100
  },
  'SF2': {
    'TAUX_REALISATION_CORRECTIF/PT':82.1,'OT préparation <1 mois':74.5,'OT préparation >3 mois':7.8,
    'OT préparation 1mois< <3mois':17.7,'OT planification <1 mois':79.2,'OT planification >3 mois':6.1,
    'OT planification 1mois< <3mois':14.7,'OT exécution <1 mois':81.3,'OT exécution >3 mois':4.2,
    'OT exécution 1mois< <3mois':14.5,'Performance Graissage':91.5,'Performance Inspection':93.2,
    'Performance Appels Systématiques':89.7,'appel avis approuvé':92.3,'OT LANC ESTIME':94.1,
    'Backlog préparation caractérisé':88.5,'Backlog planification caractérisé':91.2,
    'OT CONFIME':96.3,'OT_COR_EGAL':93.7,'OT Fiabilité':100,'Total Avis de Panne':100
  },
  'UTA': {
    'TAUX_REALISATION_CORRECTIF/PT':79.2,'OT préparation <1 mois':68.3,'OT préparation >3 mois':11.2,
    'OT préparation 1mois< <3mois':20.5,'OT planification <1 mois':72.1,'OT planification >3 mois':9.3,
    'OT planification 1mois< <3mois':18.6,'OT exécution <1 mois':76.8,'OT exécution >3 mois':8.1,
    'OT exécution 1mois< <3mois':15.1,'Performance Graissage':88.4,'Performance Inspection':90.1,
    'Performance Appels Systématiques':85.3,'appel avis approuvé':98.1,'OT LANC ESTIME':100,
    'Backlog préparation caractérisé':100,'Backlog planification caractérisé':99.2,
    'OT CONFIME':100,'OT_COR_EGAL':98.5,'OT Fiabilité':100,'Total Avis de Panne':100
  },
  'UTB': {
    'TAUX_REALISATION_CORRECTIF/PT':91.3,'OT préparation <1 mois':88.1,'OT préparation >3 mois':2.1,
    'OT préparation 1mois< <3mois':9.8,'OT planification <1 mois':90.5,'OT planification >3 mois':1.5,
    'OT planification 1mois< <3mois':8.0,'OT exécution <1 mois':85.2,'OT exécution >3 mois':3.8,
    'OT exécution 1mois< <3mois':11.0,'Performance Graissage':97.8,'Performance Inspection':96.5,
    'Performance Appels Systématiques':98.2,'appel avis approuvé':94.2,'OT LANC ESTIME':96.5,
    'Backlog préparation caractérisé':94.8,'Backlog planification caractérisé':93.1,
    'OT CONFIME':97.2,'OT_COR_EGAL':91.3,'OT Fiabilité':100,'Total Avis de Panne':100
  },
  'ATC': {
    'TAUX_REALISATION_CORRECTIF/PT':72.5,'OT préparation <1 mois':61.2,'OT préparation >3 mois':15.3,
    'OT préparation 1mois< <3mois':23.5,'OT planification <1 mois':65.8,'OT planification >3 mois':12.7,
    'OT planification 1mois< <3mois':21.5,'OT exécution <1 mois':70.1,'OT exécution >3 mois':10.4,
    'OT exécution 1mois< <3mois':19.5,'Performance Graissage':82.1,'Performance Inspection':84.6,
    'Performance Appels Systématiques':78.9,'appel avis approuvé':87.5,'OT LANC ESTIME':89.3,
    'Backlog préparation caractérisé':82.1,'Backlog planification caractérisé':85.4,
    'OT CONFIME':91.8,'OT_COR_EGAL':88.2,'OT Fiabilité':100,'Total Avis de Panne':100
  },
  'ATD': {
    'TAUX_REALISATION_CORRECTIF/PT':85.8,'OT préparation <1 mois':79.5,'OT préparation >3 mois':5.1,
    'OT préparation 1mois< <3mois':15.4,'OT planification <1 mois':82.3,'OT planification >3 mois':4.8,
    'OT planification 1mois< <3mois':12.9,'OT exécution <1 mois':80.5,'OT exécution >3 mois':5.5,
    'OT exécution 1mois< <3mois':14.0,'Performance Graissage':93.7,'Performance Inspection':95.1,
    'Performance Appels Systématiques':92.4,'appel avis approuvé':95.8,'OT LANC ESTIME':97.8,
    'Backlog préparation caractérisé':96.2,'Backlog planification caractérisé':95.5,
    'OT CONFIME':98.4,'OT_COR_EGAL':96.1,'OT Fiabilité':100,'Total Avis de Panne':100
  }
};

// =====================================================================
// DONNÉES BACKLOG
// =====================================================================
const OMS_DATA = {
  'SF1':{'CRÉÉ':12,'LANC':28,'CLOT':45,'TCLO':8},
  'SF2':{'CRÉÉ':18,'LANC':32,'CLOT':38,'TCLO':6},
  'UTA':{'CRÉÉ':25,'LANC':35,'CLOT':30,'TCLO':5},
  'UTB':{'CRÉÉ':8,'LANC':22,'CLOT':52,'TCLO':12},
  'ATC':{'CRÉÉ':30,'LANC':40,'CLOT':20,'TCLO':4},
  'ATD':{'CRÉÉ':15,'LANC':26,'CLOT':40,'TCLO':10}
};
const THERMO_DATA = {
  'SF1':{'CRÉÉ':3,'LANC':7,'CLOT':12,'TCLO':2},
  'SF2':{'CRÉÉ':5,'LANC':8,'CLOT':9,'TCLO':1},
  'UTA':{'CRÉÉ':6,'LANC':9,'CLOT':8,'TCLO':1},
  'UTB':{'CRÉÉ':2,'LANC':5,'CLOT':14,'TCLO':3},
  'ATC':{'CRÉÉ':8,'LANC':10,'CLOT':5,'TCLO':1},
  'ATD':{'CRÉÉ':4,'LANC':6,'CLOT':10,'TCLO':2}
};
const BACKLOG_PREP = {
  'SF1':{'<1 mois':8,'1-3 mois':3,'>3 mois':1,'Non caractérisé':0},
  'SF2':{'<1 mois':10,'1-3 mois':5,'>3 mois':3,'Non caractérisé':2},
  'UTA':{'<1 mois':12,'1-3 mois':8,'>3 mois':4,'Non caractérisé':6},
  'UTB':{'<1 mois':5,'1-3 mois':2,'>3 mois':1,'Non caractérisé':0},
  'ATC':{'<1 mois':10,'1-3 mois':10,'>3 mois':7,'Non caractérisé':8},
  'ATD':{'<1 mois':8,'1-3 mois':4,'>3 mois':2,'Non caractérisé':1}
};
const BACKLOG_PLAN = {
  'SF1':{'<1 mois':18,'1-3 mois':7,'>3 mois':3,'Non caractérisé':0},
  'SF2':{'<1 mois':16,'1-3 mois':10,'>3 mois':5,'Non caractérisé':3},
  'UTA':{'<1 mois':14,'1-3 mois':12,'>3 mois':6,'Non caractérisé':7},
  'UTB':{'<1 mois':14,'1-3 mois':5,'>3 mois':2,'Non caractérisé':1},
  'ATC':{'<1 mois':12,'1-3 mois':14,'>3 mois':10,'Non caractérisé':10},
  'ATD':{'<1 mois':14,'1-3 mois':8,'>3 mois':3,'Non caractérisé':1}
};
const BACKLOG_EXEC = {
  'SF1':{'<1 mois':20,'1-3 mois':12,'>3 mois':4},
  'SF2':{'<1 mois':18,'1-3 mois':16,'>3 mois':6},
  'UTA':{'<1 mois':15,'1-3 mois':18,'>3 mois':8},
  'UTB':{'<1 mois':16,'1-3 mois':8,'>3 mois':3},
  'ATC':{'<1 mois':12,'1-3 mois':20,'>3 mois':12},
  'ATD':{'<1 mois':18,'1-3 mois':10,'>3 mois':4}
};

// =====================================================================
// ACTIONS RECOMMANDÉES
// =====================================================================
const ACTIONS_MAP = {
  'TAUX_REALISATION_CORRECTIF/PT':"Améliorer le taux de réalisation des OT correctifs par poste de travail.",
  'OT préparation <1 mois':"Réduire l'âge de préparation des OT (< 1 mois). Accélérer le traitement CRPR.",
  'OT préparation >3 mois':"Traiter en priorité les OT avec préparation > 3 mois. Planifier des sessions dédiées.",
  'OT préparation 1mois< <3mois':"Réduire les OT entre 1 et 3 mois de préparation. Suivi hebdomadaire.",
  'OT planification <1 mois':"Réduire l'âge de planification des OT (< 1 mois). Optimiser le processus ATPL.",
  'OT planification >3 mois':"Traiter les OT avec planification > 3 mois. Revue de backlog planification.",
  'OT planification 1mois< <3mois':"Réduire les OT entre 1 et 3 mois de planification.",
  'OT exécution <1 mois':"Réduire l'âge d'exécution des OT (< 1 mois). Anticiper les ressources.",
  'OT exécution >3 mois':"Traiter en urgence les OT avec exécution > 3 mois.",
  'OT exécution 1mois< <3mois':"Réduire les OT entre 1 et 3 mois d'exécution.",
  'Performance Graissage':"Améliorer le taux de réalisation des OT de graissage (Type 350). Planifier les tournées.",
  'Performance Inspection':"Améliorer le taux de réalisation des OT d'inspection (Types 290, 300, 310).",
  'Performance Appels Systématiques':"Améliorer le taux de réalisation des appels systématiques (Type 360).",
  'appel avis approuvé':"Créer un OT pour chaque avis sans ordre approuvé (APRV).",
  'OT LANC ESTIME':"Estimer les coûts des OT lancés sans budget. Rendre le champ obligatoire.",
  'Backlog préparation caractérisé':"Caractériser le backlog de préparation (CRPR). Attribuer les statuts manquants.",
  'Backlog planification caractérisé':"Caractériser le backlog de planification (ATPL). Attribuer les statuts manquants.",
  'OT CONFIME':"Confirmer les OT terminés dont le statut système n'est pas CONF.",
  'OT_COR_EGAL':"Rapprocher les coûts réels et budgétés. Investiguer les écarts significatifs.",
  'OT Fiabilité':"Maintenir la fiabilité des OT à 100%. Aucune donnée manquante acceptée.",
  'Total Avis de Panne':"Maintenir le suivi des avis de panne à 100%. Traiter tous les avis reçus."
};

// =====================================================================
// CALCULS
// =====================================================================
function calcScore(poste, kpiList) {
  let met = 0;
  kpiList.forEach(k => { if (gscoreMet(k, poste[k])) met++; });
  return kpiList.length > 0 ? (met / kpiList.length * 100) : 100;
}
function calcGroupScore(postes, kpiList) {
  if (!postes.length) return 0;
  return postes.reduce((s, p) => s + calcScore(DATA[p], kpiList), 0) / postes.length;
}
function isAnomaly(k, v) { return !gscoreMet(k, parseFloat(v)); }

// =====================================================================
// COMPOSANTS HTML
// =====================================================================
function kpiCard(type, val, lbl, icon) {
  return '<div class="kpi-card ' + type + '">' +
    '<div class="val"><span class="icon"><i class="fas ' + icon + '"></i></span>' + val + '</div>' +
    '<div class="lbl">' + lbl + '</div></div>';
}

function miniBarRow(pct, color, label) {
  const w = Math.max(0, Math.min(pct, 100));
  return '<div class="mini-bar-row">' +
    '<span class="bar-label">' + label + '</span>' +
    '<div class="pbar"><div class="pbar-fill" style="width:' + w + '%;background:' + color + '"></div></div>' +
    '<span class="bar-val" style="color:' + color + '">' + fmt(pct) + '%</span></div>';
}

function renderMiniBars(postes) {
  let h = '<div class="mini-bars">';
  postes.forEach(p => {
    const ps = calcScore(DATA[p], PERF_KPI);
    const qs = calcScore(DATA[p], QUAL_KPI);
    h += '<div class="mini-bar-item"><span class="poste-name">' + p + '</span><div class="bars-col">' +
      miniBarRow(ps, 'var(--perf)', 'P') +
      miniBarRow(qs, 'var(--qual)', 'Q') +
      '</div></div>';
  });
  h += '</div>';
  return h;
}

// =====================================================================
// TABLEAU KPI GÉNÉRIQUE
// =====================================================================
function renderKpiTable(kpiList, postes, tableId, scoreLabel, headClass) {
  let h = '<div class="tbl-wrap"><table class="kpi-tbl" id="' + tableId + '">';
  h += '<thead><tr><th class="fc-head">Poste de travail</th>';
  kpiList.forEach(k => { h += '<th>' + k + '</th>'; });
  h += '<th class="sc-head">' + scoreLabel + '</th></tr></thead><tbody>';
  let scoreSum = 0;
  postes.forEach(p => {
    const d = DATA[p]; if (!d) return;
    const sc = calcScore(d, kpiList); scoreSum += sc;
    h += '<tr><td class="fc">' + p + '</td>';
    kpiList.forEach(k => { h += '<td class="' + getCellClass(k, d[k]) + '">' + fmt(d[k]) + '%</td>'; });
    h += '<td class="sc ' + getScoreClass(sc) + '">' + fmt(sc) + '%</td></tr>';
  });
  // Total Général
  const avgSc = postes.length > 0 ? scoreSum / postes.length : 0;
  h += '<tr class="total-row"><td class="fc">Total Général</td>';
  kpiList.forEach(k => {
    const avg = postes.reduce((s, p) => s + (DATA[p] ? DATA[p][k] : 0), 0) / postes.length;
    h += '<td class="' + getCellClass(k, avg) + '">' + fmt(avg) + '%</td>';
  });
  h += '<td class="sc ' + getScoreClass(avgSc) + '">' + fmt(avgSc) + '%</td></tr>';
  h += '</tbody></table></div>';
  return h;
}

// =====================================================================
// PAGE DASHBOARD
// =====================================================================
function renderDashboard() {
  const el = document.getElementById('page-dashboard');
  let h = '';
  // En-tête
  h += '<div class="page-header"><h1><i class="fas fa-tachometer-alt"></i>Performance & Qualité</h1>' +
    '<span class="badge"><i class="fas fa-calendar-alt"></i>' + DATE_STR + '</span></div>';

  // Cartes résumé global
  const totalOms = ALL_POSTES.reduce((s, p) => {
    const o = OMS_DATA[p]; return s + o.CRÉÉ + o.LANC + o.CLOT + o.TCLO;
  }, 0);
  const globalPerf = calcGroupScore(ALL_POSTES, PERF_KPI);
  const globalQual = calcGroupScore(ALL_POSTES, QUAL_KPI);
  const globalTaux = ALL_POSTES.reduce((s, p) => s + DATA[p]['TAUX_REALISATION_CORRECTIF/PT'], 0) / ALL_POSTES.length;

  h += '<div class="cards-row cards-4 fade-in">';
  h += kpiCard('neutral', totalOms, 'Total OT (OMS)', 'fa-clipboard-list');
  h += kpiCard('neutral', fmt(globalTaux) + '%', 'Taux Réalisation Moyen', 'fa-percentage');
  h += kpiCard('perf', fmt(globalPerf) + '%', 'Score Performance Global', 'fa-tachometer-alt');
  h += kpiCard('qual', fmt(globalQual) + '%', 'Score Qualité Global', 'fa-check-circle');
  h += '</div>';

  // Groupes : SF1, SF2, Autres
  const groupConfigs = [
    { key: 'sf1', label: 'SF1', posts: GROUPS.sf1 },
    { key: 'sf2', label: 'SF2', posts: GROUPS.sf2 },
    { key: 'autres', label: 'Autres Postes', posts: GROUPS.autres }
  ];
  groupConfigs.forEach((g, gi) => {
    const posts = g.posts;
    const pSc = calcGroupScore(posts, PERF_KPI);
    const qSc = calcGroupScore(posts, QUAL_KPI);
    const delay = Math.min(gi + 1, 4);

    h += '<div class="group-label fd' + delay + '"><span class="dot"></span>' + g.label +
      ' <span class="posts-list">' + posts.join(' / ') + '</span></div>';

    h += '<div class="cards-row cards-2 fd' + delay + '">';
    h += kpiCard('perf', fmt(pSc) + '%', 'Score Performance — ' + g.label, 'fa-tachometer-alt');
    h += kpiCard('qual', fmt(qSc) + '%', 'Score Qualité — ' + g.label, 'fa-check-circle');
    h += '</div>';

    // Mini-barres par poste du groupe
    if (posts.length > 1) {
      h += '<div class="chart-box fd' + delay + '" style="margin-bottom:12px">' +
        '<div class="chart-title"><i class="fas fa-chart-bar"></i>Détail par poste — ' + g.label + '</div>';
      h += renderMiniBars(posts);
      h += '</div>';
    }
  });

  // Tableau Performance
  h += '<div class="section-title perf-t fd1"><i class="fas fa-tachometer-alt"></i>Indicateurs de Performance</div>';
  h += renderKpiTable(PERF_KPI, ALL_POSTES, 'tbl-perf', 'Score Perf');

  // Tableau Qualité
  h += '<div class="section-title qual-t fd2"><i class="fas fa-check-circle"></i>Indicateurs de Qualité</div>';
  h += renderKpiTable(QUAL_KPI, ALL_POSTES, 'tbl-qual', 'Score Qual');

  // Actions Recommandées
  h += '<div class="section-title ano-t fd3"><i class="fas fa-lightbulb"></i>Actions Recommandées</div>';
  h += renderActions();

  el.innerHTML = h;
}

// =====================================================================
// ACTIONS RECOMMANDÉES
// =====================================================================
function renderActions() {
  let rows = [];
  ALL_KPI.forEach(k => {
    const avg = ALL_POSTES.reduce((s, p) => s + DATA[p][k], 0) / ALL_POSTES.length;
    const cib = CIBLES[k] || 100;
    const met = LOWER_BETTER.includes(k) ? avg <= cib : avg >= cib;
    if (!met) rows.push({ kpi: k, val: avg, cib: cib, met: false });
  });
  if (!rows.length) {
    return '<div class="chart-box"><div class="chart-title"><i class="fas fa-check-circle"></i>Aucune action requise</div>' +
      '<p style="color:#059669;font-weight:600;padding:12px">Tous les indicateurs atteignent leurs objectifs à ce jour.</p></div>';
  }
  // Trier par écart décroissant
  rows.sort((a, b) => Math.abs(b.val - b.cib) - Math.abs(a.val - a.cib));
  let h = '<div class="tbl-wrap"><table class="kpi-tbl"><thead><tr>' +
    '<th class="fc-head act-head" style="min-width:260px">KPI</th>' +
    '<th class="act-head">Valeur Actuelle</th><th class="act-head">Cible</th>' +
    '<th class="act-head">Écart</th><th class="act-head">Statut</th>' +
    '<th class="act-head" style="min-width:300px">Action Recommandée</th></tr></thead><tbody>';
  rows.forEach(r => {
    const diff = r.val - r.cib;
    const ecClr = diff > 0 ? '#dc2626' : '#059669';
    h += '<tr><td class="fc" style="background:#fef2f2!important;color:#991b1b!important">' + r.kpi + '</td>' +
      '<td>' + fmt(r.val) + '%</td><td>' + r.cib + '%</td>' +
      '<td style="color:' + ecClr + ';font-weight:700">' + (diff > 0 ? '+' : '') + fmt(diff) + '%</td>' +
      '<td class="cb">NON ATTEINT</td>' +
      '<td style="text-align:left;color:#475569;font-size:11px;white-space:normal;max-width:320px">' + (ACTIONS_MAP[r.kpi] || '') + '</td></tr>';
  });
  h += '</tbody></table></div>';
  return h;
}

// =====================================================================
// PAGE ANOMALIES
// =====================================================================
function renderAnomalies() {
  const el = document.getElementById('page-anomalies');
  let h = '';
  h += '<div class="page-header"><h1><i class="fas fa-exclamation-triangle"></i>Anomalies</h1>' +
    '<span class="badge"><i class="fas fa-calendar-alt"></i>' + DATE_STR + '</span></div>';

  // Calcul des anomalies par KPI x Poste
  let anoByKpi = {};
  ALL_KPI.forEach(k => {
    anoByKpi[k] = {};
    ALL_POSTES.forEach(p => {
      // Les indicateurs de fiabilité basés sur les avis affichent 0 si aucune donnée
      if (k === 'OT Fiabilité' || k === 'Total Avis de Panne') {
        anoByKpi[k][p] = 0; // Toujours 0, pas de données manquantes
      } else {
        anoByKpi[k][p] = isAnomaly(k, DATA[p][k]) ? 1 : 0;
      }
    });
  });

  // Anomalies par poste
  let anoByPoste = {};
  ALL_POSTES.forEach(p => {
    anoByPoste[p] = 0;
    ALL_KPI.forEach(k => { anoByPoste[p] += anoByKpi[k][p]; });
  });

  const totalAno = ALL_POSTES.reduce((s, p) => s + anoByPoste[p], 0);
  // KPI le plus problématique
  let worstKpi = ALL_KPI[0], worstKpiCount = 0;
  ALL_KPI.forEach(k => {
    const cnt = ALL_POSTES.reduce((s, p) => s + anoByKpi[k][p], 0);
    if (cnt > worstKpiCount) { worstKpiCount = cnt; worstKpi = k; }
  });
  // Poste le plus problématique
  let worstPoste = ALL_POSTES[0];
  ALL_POSTES.forEach(p => { if (anoByPoste[p] > anoByPoste[worstPoste]) worstPoste = p; });
  // KPIs sans anomalie
  const kpisOk = ALL_KPI.filter(k => ALL_POSTES.every(p => anoByKpi[k][p] === 0)).length;

  h += '<div class="cards-row cards-4 fade-in">';
  h += kpiCard('ano', totalAno, 'Total Anomalies', 'fa-bug');
  h += kpiCard('ano', worstKpiCount, 'Pire KPI : ' + worstKpi.substring(0, 22), 'fa-arrow-up');
  h += kpiCard('ano', anoByPoste[worstPoste], 'Pire Poste : ' + worstPoste, 'fa-map-marker-alt');
  h += kpiCard('neutral', kpisOk + '/' + ALL_KPI.length, 'KPIs sans anomalie', 'fa-check');
  h += '</div>';

  // Tableau Performance — classement par anomalies décroissant
  h += '<div class="section-title ano-t fd1"><i class="fas fa-sort-amount-down"></i>Classement des Anomalies — Performance</div>';
  h += renderAnoTable(PERF_KPI, anoByKpi, 'ano-perf');

  // Tableau Qualité — classement par anomalies décroissant
  h += '<div class="section-title ano-t fd2"><i class="fas fa-sort-amount-down"></i>Classement des Anomalies — Qualité</div>';
  h += renderAnoTable(QUAL_KPI, anoByKpi, 'ano-qual');

  // Résumé par poste
  h += '<div class="section-title ano-t fd3"><i class="fas fa-map-marker-alt"></i>Anomalies par Poste de Travail</div>';
  h += renderAnoByPoste(anoByPoste, anoByKpi);

  el.innerHTML = h;
}

function renderAnoTable(kpiList, anoByKpi, tableId) {
  // Trier par nombre d'anomalies décroissant
  const sorted = [...kpiList].sort((a, b) => {
    const sa = ALL_POSTES.reduce((s, p) => s + anoByKpi[a][p], 0);
    const sb = ALL_POSTES.reduce((s, p) => s + anoByKpi[b][p], 0);
    return sb - sa;
  });
  let h = '<div class="tbl-wrap"><table class="kpi-tbl" id="' + tableId + '">';
  h += '<thead><tr><th class="fc-head ano-head" style="min-width:300px">Indicateur KPI</th>';
  ALL_POSTES.forEach(p => { h += '<th class="ano-head">' + p + '</th>'; });
  h += '<th class="sc-head ano-head">Total</th></tr></thead><tbody>';
  sorted.forEach(k => {
    let tot = 0;
    h += '<tr><td class="fc" style="white-space:normal;max-width:300px;background:#fef2f2!important;color:#7f1d1d!important">' + k + '</td>';
    ALL_POSTES.forEach(p => {
      const n = anoByKpi[k][p]; tot += n;
      h += '<td class="' + getAnoCellClass(n) + '" style="text-align:center">' + n + '</td>';
    });
    h += '<td class="sc ' + getAnoCellClass(tot) + '" style="text-align:center;font-weight:800">' + tot + '</td></tr>';
  });
  // Total Général
  h += '<tr class="total-row"><td class="fc">Total Général</td>';
  let gTot = 0;
  ALL_POSTES.forEach(p => {
    let s = 0;
    kpiList.forEach(k => { s += anoByKpi[k][p]; });
    gTot += s;
    h += '<td style="text-align:center;font-weight:800">' + s + '</td>';
  });
  h += '<td class="sc" style="text-align:center;font-weight:800">' + gTot + '</td></tr>';
  h += '</tbody></table></div>';
  return h;
}

function renderAnoByPoste(anoByPoste, anoByKpi) {
  const sorted = [...ALL_POSTES].sort((a, b) => anoByPoste[b] - anoByPoste[a]);
  let h = '<div class="tbl-wrap"><table class="kpi-tbl"><thead><tr>' +
    '<th class="fc-head ano-head">Poste de travail</th>' +
    '<th class="ano-head">Nb Anomalies</th>' +
    '<th class="ano-head" style="min-width:400px">Indicateurs en anomalie</th>' +
    '<th class="sc-head ano-head">Sévérité</th></tr></thead><tbody>';
  sorted.forEach(p => {
    const n = anoByPoste[p];
    const kpis = ALL_KPI.filter(k => anoByKpi[k][p] === 1);
    const sev = n === 0 ? 'Aucune' : (n <= 3 ? 'Faible' : (n <= 6 ? 'Modérée' : 'Critique'));
    const sevCls = n === 0 ? 'cg' : (n <= 3 ? 'cw' : (n <= 6 ? 'a2' : 'a3'));
    h += '<tr><td class="fc">' + p + '</td>' +
      '<td class="' + getAnoCellClass(n) + '" style="text-align:center;font-weight:800">' + n + '</td>' +
      '<td style="text-align:left;font-size:11px;color:#475569;white-space:normal;max-width:420px">' +
      (kpis.length ? kpis.join(' &bull; ') : '<span style="color:#059669;font-weight:600">Aucun indicateur en anomalie</span>') +
      '</td>' +
      '<td class="sc ' + sevCls + '" style="text-align:center">' + sev + '</td></tr>';
  });
  h += '</tbody></table></div>';
  return h;
}

// =====================================================================
// PAGE BACKLOG
// =====================================================================
function renderBacklog() {
  const el = document.getElementById('page-backlog');
  let h = '';
  h += '<div class="page-header"><h1><i class="fas fa-layer-group"></i>Backlog</h1>' +
    '<span class="badge"><i class="fas fa-calendar-alt"></i>' + DATE_STR + '</span></div>';

  // Cartes résumé
  const totalOms = ALL_POSTES.reduce((s, p) => { const o = OMS_DATA[p]; return s + o.CRÉÉ + o.LANC + o.CLOT + o.TCLO; }, 0);
  const totalThermo = ALL_POSTES.reduce((s, p) => { const o = THERMO_DATA[p]; return s + o.CRÉÉ + o.LANC + o.CLOT + o.TCLO; }, 0);
  const totalCree = ALL_POSTES.reduce((s, p) => s + OMS_DATA[p].CRÉÉ + THERMO_DATA[p].CRÉÉ, 0);
  const totalLanc = ALL_POSTES.reduce((s, p) => s + OMS_DATA[p].LANC + THERMO_DATA[p].LANC, 0);

  h += '<div class="cards-row cards-4 fade-in">';
  h += kpiCard('neutral', totalOms, 'Total OT OMS', 'fa-clipboard-list');
  h += kpiCard('neutral', totalThermo, 'Total OT Thermographie', 'fa-temperature-high');
  h += kpiCard('bl', totalCree, 'OT en Statut CRÉÉ', 'fa-plus-circle');
  h += kpiCard('bl', totalLanc, 'OT en Statut LANC', 'fa-play-circle');
  h += '</div>';

  // OMS + Thermographie côte à côte
  h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px" class="fd1">';
  h += '<div>' + renderStatutTable(OMS_DATA, 'OT OMS par Poste et Statut', 'bl-head') +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px">' +
    '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-pie"></i>OMS — Par Statut</div><canvas id="chart-oms-statut"></canvas></div>' +
    '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-pie"></i>OMS — Réalisés vs Non Réalisés</div><canvas id="chart-oms-real"></canvas></div>' +
    '</div></div>';
  h += '<div>' + renderStatutTable(THERMO_DATA, 'OT Thermographie par Poste et Statut', 'bl-head') +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px">' +
    '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-pie"></i>Thermo — Par Statut</div><canvas id="chart-thermo-statut"></canvas></div>' +
    '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-pie"></i>Thermo — Réalisés vs Non Réalisés</div><canvas id="chart-thermo-real"></canvas></div>' +
    '</div></div>';
  h += '</div>';

  // Statistiques globales des statuts OT
  h += '<div class="section-title bl-t fd2"><i class="fas fa-chart-bar"></i>Statistiques des Statuts OT — Tous Types Confondus</div>';
  const allStats = {};
  ALL_POSTES.forEach(p => {
    allStats[p] = {
      'CRÉÉ': OMS_DATA[p].CRÉÉ + THERMO_DATA[p].CRÉÉ,
      'LANC': OMS_DATA[p].LANC + THERMO_DATA[p].LANC,
      'CLOT': OMS_DATA[p].CLOT + THERMO_DATA[p].CLOT,
      'TCLO': OMS_DATA[p].TCLO + THERMO_DATA[p].TCLO
    };
  });
  h += '<div style="display:grid;grid-template-columns:2fr 1fr;gap:14px" class="fd2">';
  h += '<div>' + renderStatutTable(allStats, 'Tous les OT par Poste et Statut', 'bl-head') + '</div>';
  h += '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-pie"></i>Répartition Globale</div><canvas id="chart-all-statut"></canvas></div>';
  h += '</div>';

  // Caractérisation Backlog
  h += '<div class="section-title bl-t fd3"><i class="fas fa-tags"></i>Caractérisation du Backlog</div>';
  h += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px" class="fd3">';
  h += '<div>' + renderBacklogTable(BACKLOG_PREP, 'Backlog Préparation', 'bl-head') +
    '<div class="chart-box" style="margin-top:8px"><div class="chart-title"><i class="fas fa-chart-bar"></i>Préparation — par Âge</div><canvas id="chart-bk-prep"></canvas></div></div>';
  h += '<div>' + renderBacklogTable(BACKLOG_PLAN, 'Backlog Planification', 'bl-head') +
    '<div class="chart-box" style="margin-top:8px"><div class="chart-title"><i class="fas fa-chart-bar"></i>Planification — par Âge</div><canvas id="chart-bk-plan"></canvas></div></div>';
  h += '<div>' + renderBacklogTable(BACKLOG_EXEC, 'Backlog Exécution', 'bl-head') +
    '<div class="chart-box" style="margin-top:8px"><div class="chart-title"><i class="fas fa-chart-bar"></i>Exécution — par Âge</div><canvas id="chart-bk-exec"></canvas></div></div>';
  h += '</div>';

  // Graphiques de synthèse
  h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px" class="fd4">';
  h += '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-bar"></i>Synthèse Backlog — Comparaison par Poste</div><canvas id="chart-bk-synth"></canvas></div>';
  h += '<div class="chart-box"><div class="chart-title"><i class="fas fa-chart-bar"></i>Synthèse Backlog — Taux de Caractérisation</div><canvas id="chart-bk-caract"></canvas></div>';
  h += '</div>';

  el.innerHTML = h;
  // Créer les graphiques après injection DOM
  setTimeout(() => createBacklogCharts(), 60);
}

function renderStatutTable(data, title, headCls) {
  const statuses = ['CRÉÉ', 'LANC', 'CLOT', 'TCLO'];
  let h = '<div class="chart-box"><div class="chart-title"><i class="fas fa-table"></i>' + title + '</div><div class="tbl-wrap"><table class="kpi-tbl">';
  h += '<thead><tr><th class="fc-head ' + headCls + '">Poste</th>';
  statuses.forEach(s => { h += '<th class="' + headCls + '">' + s + '</th>'; });
  h += '<th class="sc-head ' + headCls + '">Total</th></tr></thead><tbody>';
  let totals = {}; statuses.forEach(s => { totals[s] = 0; });
  ALL_POSTES.forEach(p => {
    const d = data[p]; if (!d) return;
    let tot = 0;
    h += '<tr><td class="fc">' + p + '</td>';
    statuses.forEach(s => {
      const v = d[s] || 0; totals[s] += v; tot += v;
      let cls = '';
      if (s === 'CRÉÉ') cls = 'cb';
      else if (s === 'LANC') cls = 'cw';
      else cls = 'cg';
      h += '<td class="' + cls + '" style="text-align:center">' + v + '</td>';
    });
    h += '<td class="sc" style="text-align:center;font-weight:800">' + tot + '</td></tr>';
  });
  let gTot = 0;
  h += '<tr class="total-row"><td class="fc">Total Général</td>';
  statuses.forEach(s => { gTot += totals[s]; h += '<td style="text-align:center;font-weight:800">' + totals[s] + '</td>'; });
  h += '<td class="sc" style="text-align:center;font-weight:800">' + gTot + '</td></tr>';
  h += '</tbody></table></div></div>';
  return h;
}

function renderBacklogTable(data, title, headCls) {
  const cats = Object.keys(data[ALL_POSTES[0]]);
  let h = '<div class="chart-box"><div class="chart-title"><i class="fas fa-table"></i>' + title + '</div><div class="tbl-wrap"><table class="kpi-tbl">';
  h += '<thead><tr><th class="fc-head ' + headCls + '">Poste</th>';
  cats.forEach(c => { h += '<th class="' + headCls + '">' + c + '</th>'; });
  h += '<th class="sc-head ' + headCls + '">Total</th></tr></thead><tbody>';
  let totals = {}; cats.forEach(c => { totals[c] = 0; });
  ALL_POSTES.forEach(p => {
    const d = data[p]; if (!d) return;
    let tot = 0;
    h += '<tr><td class="fc">' + p + '</td>';
    cats.forEach(c => {
      const v = d[c] || 0; totals[c] += v; tot += v;
      let cls = 'cg';
      if (c === '>3 mois' || c === 'Non caractérisé') cls = 'cb';
      else if (c === '1-3 mois') cls = 'cw';
      h += '<td class="' + cls + '" style="text-align:center">' + v + '</td>';
    });
    h += '<td class="sc" style="text-align:center;font-weight:800">' + tot + '</td></tr>';
  });
  let gTot = 0;
  h += '<tr class="total-row"><td class="fc">Total Général</td>';
  cats.forEach(c => { gTot += totals[c]; h += '<td style="text-align:center;font-weight:800">' + totals[c] + '</td>'; });
  h += '<td class="sc" style="text-align:center;font-weight:800">' + gTot + '</td></tr>';
  h += '</tbody></table></div></div>';
  return h;
}

// =====================================================================
// GRAPHIQUES BACKLOG (Chart.js)
// =====================================================================
const _chartInstances = [];

function destroyBacklogCharts() {
  _chartInstances.forEach(c => { try { c.destroy(); } catch (e) {} });
  _chartInstances.length = 0;
}

function createBacklogCharts() {
  destroyBacklogCharts();
  const baseOpts = {
    responsive: true, maintainAspectRatio: true,
    plugins: { legend: { labels: { font: { family: 'Space Grotesk', size: 11 }, boxWidth: 12 } } }
  };
  const pieOpts = {
    ...baseOpts,
    cutout: '52%',
    plugins: { legend: { position: 'bottom', labels: { font: { family: 'Space Grotesk', size: 10 }, boxWidth: 10, padding: 8 } } }
  };
  const statutColors = ['#dc2626', '#d97706', '#059669', '#0284c7'];
  const ageColors = ['#059669', '#d97706', '#dc2626', '#94a3b8'];

  function makePie(id, labels, values, colors) {
    const ctx = document.getElementById(id); if (!ctx) return;
    const c = new Chart(ctx, { type: 'doughnut', data: { labels, datasets: [{ data: values, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }] }, options: pieOpts });
    _chartInstances.push(c);
  }

  function makeStacked(id, data, cats, colors) {
    const ctx = document.getElementById(id); if (!ctx) return;
    const datasets = cats.map((c, i) => ({ label: c, data: ALL_POSTES.map(p => data[p][c] || 0), backgroundColor: colors[i] || colors[0], borderRadius: 2 }));
    const c = new Chart(ctx, { type: 'bar', data: { labels: ALL_POSTES, datasets }, options: { ...baseOpts, scales: { x: { stacked: true, grid: { display: false }, ticks: { font: { family: 'Space Grotesk', size: 11 } } }, y: { stacked: true, grid: { color: '#f1f5f9' }, ticks: { font: { family: 'Space Grotesk', size: 10 } } } } } });
    _chartInstances.push(c);
  }

  function aggStatuts(data) {
    const r = [0, 0, 0, 0];
    ALL_POSTES.forEach(p => { const d = data[p]; if (!d) return; r[0] += d.CRÉÉ || 0; r[1] += d.LANC || 0; r[2] += d.CLOT || 0; r[3] += d.TCLO || 0; });
    return r;
  }
  function realVsNon(data) {
    let r = 0, nr = 0;
    ALL_POSTES.forEach(p => { const d = data[p]; if (!d) return; r += (d.CLOT || 0) + (d.TCLO || 0); nr += (d.CRÉÉ || 0) + (d.LANC || 0); });
    return [r, nr];
  }

  // OMS
  makePie('chart-oms-statut', ['CRÉÉ', 'LANC', 'CLOT', 'TCLO'], aggStatuts(OMS_DATA), statutColors);
  makePie('chart-oms-real', ['Réalisés', 'Non Réalisés'], realVsNon(OMS_DATA), ['#059669', '#dc2626']);
  // Thermo
  makePie('chart-thermo-statut', ['CRÉÉ', 'LANC', 'CLOT', 'TCLO'], aggStatuts(THERMO_DATA), statutColors);
  makePie('chart-thermo-real', ['Réalisés', 'Non Réalisés'], realVsNon(THERMO_DATA), ['#059669', '#dc2626']);
  // Tous
  const allAgg = [0, 0, 0, 0];
  ALL_POSTES.forEach(p => { allAgg[0] += allStatsGlobal[p]['CRÉÉ']; allAgg[1] += allStatsGlobal[p]['LANC']; allAgg[2] += allStatsGlobal[p]['CLOT']; allAgg[3] += allStatsGlobal[p]['TCLO']; });
  makePie('chart-all-statut', ['CRÉÉ', 'LANC', 'CLOT', 'TCLO'], allAgg, statutColors);

  // Backlogs empilés
  makeStacked('chart-bk-prep', BACKLOG_PREP, ['<1 mois', '1-3 mois', '>3 mois', 'Non caractérisé'], ageColors);
  makeStacked('chart-bk-plan', BACKLOG_PLAN, ['<1 mois', '1-3 mois', '>3 mois', 'Non caractérisé'], ageColors);
  makeStacked('chart-bk-exec', BACKLOG_EXEC, ['<1 mois', '1-3 mois', '>3 mois'], ageColors.slice(0, 3));

  // Synthèse comparaison grouped bar
  const prepTots = ALL_POSTES.map(p => Object.values(BACKLOG_PREP[p]).reduce((a, b) => a + b, 0));
  const planTots = ALL_POSTES.map(p => Object.values(BACKLOG_PLAN[p]).reduce((a, b) => a + b, 0));
  const execTots = ALL_POSTES.map(p => Object.values(BACKLOG_EXEC[p]).reduce((a, b) => a + b, 0));
  const ctxS = document.getElementById('chart-bk-synth');
  if (ctxS) {
    const c = new Chart(ctxS, { type: 'bar', data: { labels: ALL_POSTES, datasets: [
      { label: 'Préparation', data: prepTots, backgroundColor: '#0d9488', borderRadius: 3 },
      { label: 'Planification', data: planTots, backgroundColor: '#7c3aed', borderRadius: 3 },
      { label: 'Exécution', data: execTots, backgroundColor: '#d97706', borderRadius: 3 }
    ]}, options: { ...baseOpts, scales: { x: { grid: { display: false }, ticks: { font: { family: 'Space Grotesk', size: 11 } } }, y: { grid: { color: '#f1f5f9' }, ticks: { font: { family: 'Space Grotesk', size: 10 } } } } } });
    _chartInstances.push(c);
  }

  // Taux de caractérisation
  const prepCaract = ALL_POSTES.map(p => { const d = BACKLOG_PREP[p]; const tot = Object.values(d).reduce((a, b) => a + b, 0); const nc = d['Non caractérisé'] || 0; return tot > 0 ? ((tot - nc) / tot * 100) : 100; });
  const planCaract = ALL_POSTES.map(p => { const d = BACKLOG_PLAN[p]; const tot = Object.values(d).reduce((a, b) => a + b, 0); const nc = d['Non caractérisé'] || 0; return tot > 0 ? ((tot - nc) / tot * 100) : 100; });
  const ctxC = document.getElementById('chart-bk-caract');
  if (ctxC) {
    const c = new Chart(ctxC, { type: 'bar', data: { labels: ALL_POSTES, datasets: [
      { label: 'Préparation %', data: prepCaract, backgroundColor: '#0d9488', borderRadius: 3 },
      { label: 'Planification %', data: planCaract, backgroundColor: '#7c3aed', borderRadius: 3 }
    ]}, options: { ...baseOpts, scales: { x: { grid: { display: false }, ticks: { font: { family: 'Space Grotesk', size: 11 } } }, y: { min: 0, max: 100, grid: { color: '#f1f5f9' }, ticks: { font: { family: 'Space Grotesk', size: 10 }, callback: v => v + '%' } } } } });
    _chartInstances.push(c);
  }
}

// Pré-calcul des stats globales pour les graphiques
const allStatsGlobal = {};
ALL_POSTES.forEach(p => {
  allStatsGlobal[p] = {
    'CRÉÉ': OMS_DATA[p].CRÉÉ + THERMO_DATA[p].CRÉÉ,
    'LANC': OMS_DATA[p].LANC + THERMO_DATA[p].LANC,
    'CLOT': OMS_DATA[p].CLOT + THERMO_DATA[p].CLOT,
    'TCLO': OMS_DATA[p].TCLO + THERMO_DATA[p].TCLO
  };
});

// =====================================================================
// PAGE SUIVI (code conservé, masqué du menu)
// =====================================================================
function renderSuivi() {
  const el = document.getElementById('page-suivi');
  const changelog = [
    { ver: '2.1', date: '18/06/2025', items: [
      "Déplacement KPI Graissage/Inspection/Systématiques de Qualité vers Performance",
      "Nouveau tableau OT OMS par Poste et Statut OT avec 2 Pie charts",
      "Nouveau tableau OT Thermographie par Poste et Statut OT avec 2 Pie charts",
      "Nouveau tableau Tous les OT par Poste et Statut OT avec 2 Pie charts",
      "Page Anomalies simplifiée : résumé KPI x Poste avec coloriage",
      "Page Suivi & Évolution : synthèse entre deux dates par poste"
    ]},
    { ver: '2.0', date: '15/06/2025', items: [
      "KPI Taux réalisation correctif/PT : ajout filtre SOPL=1, numérateur CLOT+TCLO, total=0 => 100%",
      "KPI Âge backlog préparation : filtre Statut OT=CRÉÉ + Statut utilisateur contient CRPR",
      "KPI Âge backlog planification : filtre Statut OT=LANC + Statut utilisateur contient ATPL",
      "Nouveau KPI Performance Graissage (Type 350) — Seuils V/J/R",
      "Nouveau KPI Performance Inspection (Types 290,300,310) — Exclusion dates futures — Seuils V/J/R",
      "Nouveau KPI Performance Appels Systématiques (Type 360) — Exclusion dates futures — Seuils V/J/R",
      "Nouveaux KPI Qualité Appels : OT Fiabilité (100%), Total Avis de Panne (100%)",
      "Mise en place mécanisme de cache pour éviter les recalculs systématiques",
      "Activation du suivi des améliorations et évolutions (changelog)"
    ]}
  ];
  let h = '<div class="page-header"><h1><i class="fas fa-tasks"></i>Suivi des Améliorations et Évolutions</h1>' +
    '<span class="badge"><i class="fas fa-calendar-alt"></i>' + DATE_STR + '</span></div>';
  h += '<div class="chart-box" style="max-width:820px"><div class="chart-title"><i class="fas fa-history"></i>Journal des Évolutions</div>';
  h += '<div class="evol-timeline">';
  changelog.forEach(c => {
    h += '<div class="evol-item">';
    h += '<div class="evol-ver">Version ' + c.ver + '</div>';
    h += '<div class="evol-date">' + c.date + '</div>';
    c.items.forEach(item => {
      h += '<div class="evol-change">' + item + '</div>';
    });
    h += '</div>';
  });
  h += '</div></div>';
  el.innerHTML = h;
}

// =====================================================================
// NAVIGATION
// =====================================================================
function navigateTo(page) {
  // Détruire les graphiques du backlog si on quitte la page
  if (page !== 'backlog') destroyBacklogCharts();

  document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const section = document.getElementById('page-' + page);
  if (section) section.classList.add('active');
  const navItem = document.querySelector('.nav-item[data-page="' + page + '"]');
  if (navItem) navItem.classList.add('active');

  // Rendu paresseux
  if (page === 'dashboard' && !document.getElementById('page-dashboard').innerHTML) renderDashboard();
  if (page === 'anomalies' && !document.getElementById('page-anomalies').innerHTML) renderAnomalies();
  if (page === 'backlog') {
    if (!document.getElementById('page-backlog').innerHTML) renderBacklog();
    else setTimeout(() => createBacklogCharts(), 30);
  }
  if (page === 'suivi' && !document.getElementById('page-suivi').innerHTML) renderSuivi();
}

// =====================================================================
// INITIALISATION
// =====================================================================
function init() {
  renderDashboard();
  renderAnomalies();
  renderSuivi();
  // Ne pas pré-rendre le backlog (graphiques lourds) — rendu à la demande
}
init();
</script>
</body>
</html>
