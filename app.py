# Version: 0.02
# Sorter VDL B2B — Herramienta de configuración de semanas especiales
# Uso: python app.py
# Abre automáticamente http://localhost:5001

import os, sys, json, glob, shutil, threading, webbrowser, cgi
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
SCRIPT_GD  = os.path.join(BASE_DIR, 'process_parrilla.py')
SCRIPT_GANTT = os.path.join(BASE_DIR, 'gantt_1h.py')
SCRIPT_MAP   = os.path.join(BASE_DIR, 'sorter_map_por_dia.py')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

APP_VERSION = '0.02'

# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sorter VDL B2B</title>
<script src="https://cdn.tailwindcss.com"></script>
<script>tailwind.config = { darkMode: 'class' }</script>
<style>
  body { font-family: 'Trebuchet MS', Arial, sans-serif; }
  .drop-zone { border: 2px dashed #d1d5db; transition: all .15s; }
  .drop-zone.ready { border-color: #6366f1; background: #eef2ff; }
  .dark .drop-zone.ready { background: #1e1b4b22; border-color: #818cf8; }
  .log-ok   { color: #10b981; }
  .log-warn { color: #f59e0b; }
  .log-err  { color: #ef4444; }
  .log-info { color: #9ca3af; }
  .log-sep  { color: #4b5563; border-top: 1px solid #374151; margin: 4px 0; }
</style>
</head>
<body class="bg-gray-50 dark:bg-gray-950 text-gray-900 dark:text-gray-50 min-h-screen">

<!-- TOPBAR -->
<header class="sticky top-0 z-10 bg-white dark:bg-gray-900 border-b border-gray-200 dark:border-gray-800 px-6 py-3 flex items-center justify-between">
  <div class="flex items-center gap-3">
    <div class="w-7 h-7 bg-black rounded flex items-center justify-center">
      <span class="text-white text-xs font-bold">M</span>
    </div>
    <span class="text-sm font-medium">Sorter VDL B2B &middot; Semanas especiales</span>
  </div>
  <button onclick="toggleTheme()"
    class="text-lg w-8 h-8 rounded-lg border border-gray-200 dark:border-gray-700
    hover:bg-gray-100 dark:hover:bg-gray-800 flex items-center justify-center transition" id="theme-btn">🌙</button>
</header>

<main class="max-w-2xl mx-auto px-6 py-10 space-y-6">

  <!-- STEP 01: INPUTS -->
  <section>
    <h1 class="text-xs font-semibold text-gray-400 dark:text-gray-500 uppercase tracking-widest mb-4">01 &mdash; Ficheros de entrada</h1>

    <!-- Parrilla -->
    <div class="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-4 mb-3">
      <div class="flex items-center justify-between gap-4">
        <div>
          <p class="text-sm font-medium">Parrilla de salidas</p>
          <p class="text-xs text-gray-500 mt-0.5">.xlsx &mdash; hoja <code class="bg-gray-100 dark:bg-gray-800 px-1 rounded">parrilla_test_sXX</code></p>
        </div>
        <label id="zone-parrilla" class="drop-zone cursor-pointer rounded-lg px-5 py-2.5 text-center min-w-[160px] hover:border-indigo-400 transition">
          <input type="file" accept=".xlsx" class="hidden" onchange="picked(this,'parrilla')">
          <span id="lbl-parrilla" class="text-xs text-gray-500 dark:text-gray-400">Seleccionar&hellip;</span>
        </label>
      </div>
      <div class="mt-3 flex items-center gap-3 flex-wrap">
        <div class="flex items-center gap-2">
          <label class="text-xs text-gray-500">Hoja:</label>
          <input id="sheet" type="text" value="parrilla_test_s14"
            class="text-xs border border-gray-200 dark:border-gray-700 rounded-lg px-3 py-1.5 w-44
            bg-white dark:bg-gray-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
        </div>
        <div class="flex items-center gap-2">
          <label class="text-xs text-gray-500">Semana:</label>
          <input id="semana" type="text" value="S14"
            class="text-xs border border-gray-200 dark:border-gray-700 rounded-lg px-3 py-1.5 w-20
            bg-white dark:bg-gray-800 focus:outline-none focus:ring-2 focus:ring-indigo-500">
        </div>
      </div>
    </div>

    <!-- GD -->
    <div class="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-4 mb-3 flex items-center justify-between gap-4">
      <div>
        <p class="text-sm font-medium">GRUPO_DESTINOS</p>
        <p class="text-xs text-gray-500 mt-0.5">.xlsx &mdash; export DXC o fichero clásico</p>
      </div>
      <label id="zone-gd" class="drop-zone cursor-pointer rounded-lg px-5 py-2.5 text-center min-w-[160px] hover:border-indigo-400 transition">
        <input type="file" accept=".xlsx" class="hidden" onchange="picked(this,'gd')">
        <span id="lbl-gd" class="text-xs text-gray-500 dark:text-gray-400">Seleccionar&hellip;</span>
      </label>
    </div>

    <!-- Capacity -->
    <div class="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-4 mb-3 flex items-center justify-between gap-4">
      <div>
        <p class="text-sm font-medium">Capacidad de rampas</p>
        <p class="text-xs text-gray-500 mt-0.5">.csv &mdash; <code class="bg-gray-100 dark:bg-gray-800 px-1 rounded">RAMP;PALLETS</code></p>
      </div>
      <label id="zone-cap" class="drop-zone cursor-pointer rounded-lg px-5 py-2.5 text-center min-w-[160px] hover:border-indigo-400 transition">
        <input type="file" accept=".csv" class="hidden" onchange="picked(this,'cap')">
        <span id="lbl-cap" class="text-xs text-gray-500 dark:text-gray-400">Seleccionar&hellip;</span>
      </label>
    </div>

    <!-- Bloques horarios -->
    <div class="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-4 flex items-center justify-between gap-4">
      <div>
        <p class="text-sm font-medium">Bloques horarios
          <span class="ml-1.5 text-xs font-normal text-gray-400 dark:text-gray-500">(opcional)</span>
        </p>
        <p class="text-xs text-gray-500 mt-0.5">.xlsx &mdash; para generar Gantt y Sorter Map visual</p>
      </div>
      <label id="zone-bloques" class="drop-zone cursor-pointer rounded-lg px-5 py-2.5 text-center min-w-[160px] hover:border-indigo-400 transition">
        <input type="file" accept=".xlsx" class="hidden" onchange="picked(this,'bloques')">
        <span id="lbl-bloques" class="text-xs text-gray-500 dark:text-gray-400">Seleccionar&hellip;</span>
      </label>
    </div>
  </section>

  <!-- RUN -->
  <button id="run-btn" onclick="run()"
    class="w-full py-3 rounded-xl bg-black dark:bg-white text-white dark:text-black text-sm font-medium
    hover:bg-gray-800 dark:hover:bg-gray-100 disabled:opacity-40 disabled:cursor-not-allowed transition-all">
    Generar configuración
  </button>

  <!-- LOG -->
  <section id="log-sec" class="hidden">
    <div class="flex items-center justify-between mb-2">
      <h2 class="text-xs font-semibold text-gray-400 dark:text-gray-500 uppercase tracking-widest">Proceso</h2>
      <span id="log-badge" class="text-xs text-gray-400"></span>
    </div>
    <div id="log-box"
      class="bg-gray-950 rounded-xl p-4 h-56 overflow-y-auto space-y-px font-mono text-xs text-gray-300">
    </div>
  </section>

  <!-- OUTPUTS -->
  <section id="out-sec" class="hidden">
    <h2 class="text-xs font-semibold text-gray-400 dark:text-gray-500 uppercase tracking-widest mb-4">02 &mdash; Descargar resultados</h2>
    <div class="grid grid-cols-2 gap-3" id="dl-grid"></div>
    <!-- Warnings -->
    <div id="warn-box" class="hidden mt-4 bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-800 rounded-xl p-4">
      <p class="text-xs font-semibold text-amber-800 dark:text-amber-300 mb-2">Casos para revisión manual</p>
      <div id="warn-list" class="text-xs text-amber-700 dark:text-amber-400 space-y-1"></div>
    </div>
  </section>

</main>

<footer class="text-center text-xs text-gray-300 dark:text-gray-700 py-6">
  v""" + APP_VERSION + """ &middot; Mango Logística VDL B2B
</footer>

<script>
const files = {};

function initTheme() {
  const dark = localStorage.getItem('sorter_theme') === 'dark'
    || (!localStorage.getItem('sorter_theme') && matchMedia('(prefers-color-scheme:dark)').matches);
  document.documentElement.classList.toggle('dark', dark);
  document.getElementById('theme-btn').textContent = dark ? '☀️' : '🌙';
}
function toggleTheme() {
  const d = document.documentElement.classList.toggle('dark');
  localStorage.setItem('sorter_theme', d ? 'dark' : 'light');
  document.getElementById('theme-btn').textContent = d ? '☀️' : '🌙';
}
initTheme();

function picked(input, key) {
  const f = input.files[0]; if (!f) return;
  files[key] = f;
  const lbl = document.getElementById('lbl-' + key);
  lbl.textContent = f.name;
  lbl.className = 'text-xs font-medium text-indigo-600 dark:text-indigo-400';
  document.getElementById('zone-' + key).classList.add('ready');
}

function addLog(text, cls) {
  const box = document.getElementById('log-box');
  const d = document.createElement('div');
  if (cls === 'sep') {
    d.className = 'log-sep pt-1 mt-1 text-gray-600 dark:text-gray-500';
    d.textContent = text;
  } else {
    d.className = cls || (text.startsWith('✓') ? 'log-ok'
      : text.startsWith('⚠') ? 'log-warn'
      : text.startsWith('❌') ? 'log-err' : 'log-info');
    d.textContent = text;
  }
  box.appendChild(d);
  box.scrollTop = box.scrollHeight;
}

function makeDownloadCard(href, icon, title, subtitle, colorClass) {
  return `<a href="${href}"
    class="flex items-center gap-3 bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-800
    rounded-xl px-4 py-3.5 hover:border-${colorClass}-400 hover:bg-${colorClass}-50 dark:hover:bg-${colorClass}-950 transition group cursor-pointer">
    <div class="w-9 h-9 rounded-lg bg-${colorClass}-100 dark:bg-${colorClass}-900/60 flex items-center justify-center
      text-${colorClass}-700 dark:text-${colorClass}-400 text-base font-bold">${icon}</div>
    <div>
      <p class="text-sm font-medium group-hover:text-${colorClass}-700 dark:group-hover:text-${colorClass}-400">${title}</p>
      <p class="text-xs text-gray-500">${subtitle}</p>
    </div>
  </a>`;
}

async function run() {
  if (!files.parrilla || !files.gd || !files.cap) {
    alert('Selecciona los tres ficheros obligatorios antes de continuar.'); return;
  }
  const sheet  = document.getElementById('sheet').value.trim();
  const semana = document.getElementById('semana').value.trim() || sheet.toUpperCase();
  if (!sheet) { alert('Indica el nombre de la hoja.'); return; }

  const btn = document.getElementById('run-btn');
  btn.disabled = true; btn.textContent = 'Procesando…';
  document.getElementById('log-sec').classList.remove('hidden');
  document.getElementById('out-sec').classList.add('hidden');
  document.getElementById('warn-box').classList.add('hidden');
  document.getElementById('log-box').innerHTML = '';
  document.getElementById('log-badge').textContent = '';
  addLog('Subiendo ficheros y procesando…');

  const fd = new FormData();
  fd.append('parrilla', files.parrilla);
  fd.append('gd',       files.gd);
  fd.append('cap',      files.cap);
  fd.append('sheet',    sheet);
  fd.append('semana',   semana);
  if (files.bloques) fd.append('bloques', files.bloques);

  try {
    const res  = await fetch('/run', { method: 'POST', body: fd });
    const data = await res.json();

    // Separate log sections
    (data.log_gd     || []).forEach(l => addLog(l));
    if (data.log_gantt && data.log_gantt.length) {
      addLog('── Gantt 1H ──────────────────', 'sep');
      data.log_gantt.forEach(l => addLog(l));
    }
    if (data.log_map && data.log_map.length) {
      addLog('── Sorter Map por día ─────────', 'sep');
      data.log_map.forEach(l => addLog(l));
    }

    if (data.ok) {
      document.getElementById('log-badge').textContent = '✓ Completado';
      document.getElementById('log-badge').className = 'text-xs text-emerald-500';
      document.getElementById('out-sec').classList.remove('hidden');

      // Build download cards based on what was generated
      const grid = document.getElementById('dl-grid');
      grid.innerHTML = '';
      grid.innerHTML += makeDownloadCard('/download?file=xlsx', '↓', 'GRUPO_DESTINOS', '.xlsx · subir a DXC', 'emerald');
      grid.innerHTML += makeDownloadCard('/download?file=html', '↓', 'Resumen HTML', '.html · informe con gráfico', 'blue');
      if (data.has_gantt)
        grid.innerHTML += makeDownloadCard('/download?file=gantt', '↓', 'Gantt 1H', '.xlsx · visual por bloque', 'violet');
      if (data.has_map)
        grid.innerHTML += makeDownloadCard('/download?file=map', '↓', 'Sorter Map', '.xlsx · 1 pestaña por día', 'orange');

      if ((data.warnings || []).length) {
        document.getElementById('warn-box').classList.remove('hidden');
        document.getElementById('warn-list').innerHTML =
          data.warnings.map(w => `<div>${w}</div>`).join('');
      }
    } else {
      document.getElementById('log-badge').textContent = '✗ Error';
      document.getElementById('log-badge').className = 'text-xs text-red-500';
    }
  } catch(e) {
    addLog('❌ Error de conexión: ' + e.message);
  }
  btn.disabled = false; btn.textContent = 'Generar configuración';
}
</script>
</body>
</html>"""


# ── HTTP Handler ──────────────────────────────────────────────────────────────

class Handler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass

    def do_GET(self):
        path = urlparse(self.path).path
        qs   = parse_qs(urlparse(self.path).query)

        if path == '/':
            body = HTML.encode('utf-8')
            self._respond(200, 'text/html; charset=utf-8', body)

        elif path == '/download':
            ftype = qs.get('file', [''])[0]
            patterns = {
                'xlsx':  'GRUPO_DESTINOS*.xlsx',
                'html':  'resumen_sorter*.html',
                'gantt': 'gantt_1h_*.xlsx',
                'map':   'sorter_map_*.xlsx',
            }
            mimes = {
                'xlsx':  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'html':  'text/html; charset=utf-8',
                'gantt': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'map':   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            }
            pattern = patterns.get(ftype)
            if not pattern:
                self._respond(404, 'text/plain', b'Unknown file type'); return
            found = sorted(glob.glob(os.path.join(OUTPUT_DIR, pattern)))
            if not found:
                self._respond(404, 'text/plain', b'File not ready yet'); return
            fpath = found[-1]
            fname = os.path.basename(fpath)
            with open(fpath, 'rb') as f:
                data = f.read()
            self.send_response(200)
            self.send_header('Content-Type', mimes[ftype])
            self.send_header('Content-Disposition', f'attachment; filename="{fname}"')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            self._respond(404, 'text/plain', b'Not found')

    def do_POST(self):
        if self.path != '/run':
            self._respond(404, 'text/plain', b'Not found'); return

        ctype, pdict = cgi.parse_header(self.headers.get('Content-Type', ''))
        if ctype != 'multipart/form-data':
            self._respond(400, 'text/plain', b'Expected multipart/form-data'); return

        pdict['boundary'] = bytes(pdict['boundary'], 'utf-8')
        pdict['CONTENT-LENGTH'] = int(self.headers.get('Content-Length', 0))
        form = cgi.parse_multipart(self.rfile, pdict)

        def field(key):
            v = form.get(key, [b''])[0]
            return v.decode('utf-8') if isinstance(v, bytes) else (v or '')

        sheet  = field('sheet')
        semana = field('semana') or field('sheet').upper()

        # Save uploaded files
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        paths = {}
        for key, ext in [('parrilla','.xlsx'),('gd','.xlsx'),('cap','.csv')]:
            raw = form.get(key, [None])[0]
            if raw is None:
                self._json({'ok':False,'log_gd':[f'❌ Falta fichero: {key}'],'warnings':[]})
                return
            fpath = os.path.join(UPLOAD_DIR, f'{key}_{ts}{ext}')
            with open(fpath, 'wb') as f:
                f.write(raw if isinstance(raw, bytes) else raw.encode())
            paths[key] = fpath

        # Optional: bloques_horarios
        bloques_raw = form.get('bloques', [None])[0]
        if bloques_raw:
            bloques_path = os.path.join(UPLOAD_DIR, f'bloques_{ts}.xlsx')
            with open(bloques_path, 'wb') as f:
                f.write(bloques_raw if isinstance(bloques_raw, bytes) else bloques_raw.encode())
            paths['bloques'] = bloques_path

        import subprocess

        def run_script(cmd, timeout=180):
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
            out = ((r.stdout or '') + '\n' + (r.stderr or '')).strip()
            lines = [l.rstrip() for l in out.splitlines() if l.strip()]
            return r.returncode, lines

        # ── Step 1: process_parrilla ──
        code, log_gd = run_script([
            sys.executable, SCRIPT_GD,
            paths['parrilla'], paths['gd'], paths['cap'], sheet, semana,
        ])
        warnings = [l for l in log_gd if '⚠' in l or '❌' in l]

        if code != 0:
            log_gd.append(f'❌ Proceso terminó con error (código {code})')
            self._json({'ok':False,'log_gd':log_gd,'log_gantt':[],'log_map':[],'warnings':warnings})
            return

        # Find and copy GD output
        run_start = os.path.getmtime(paths['parrilla'])
        gd_xlsx = self._copy_newest(f'GRUPO_DESTINOS*.xlsx', run_start)
        self._copy_newest('resumen_sorter*.html', run_start)
        log_gd.append('✓ Ficheros GD listos.')

        # ── Step 2 & 3: visual scripts (only if bloques provided) ──
        log_gantt, log_map = [], []
        has_gantt, has_map = False, False

        if 'bloques' in paths and gd_xlsx:
            ts2 = datetime.now().strftime('%Y%m%d_%H%M%S')

            gantt_out = os.path.join(OUTPUT_DIR, f'gantt_1h_{semana}_{ts2}.xlsx')
            code_g, log_gantt = run_script([
                sys.executable, SCRIPT_GANTT,
                paths['cap'], gd_xlsx, paths['bloques'], gantt_out, 'Hoja1',
            ])
            if code_g == 0:
                has_gantt = True
                log_gantt.append('✓ Gantt 1H generado.')
            else:
                log_gantt.append(f'❌ Gantt terminó con error (código {code_g})')

            map_out = os.path.join(OUTPUT_DIR, f'sorter_map_{semana}_{ts2}.xlsx')
            code_m, log_map = run_script([
                sys.executable, SCRIPT_MAP,
                paths['cap'], gd_xlsx, paths['bloques'], map_out, 'Hoja1',
            ])
            if code_m == 0:
                has_map = True
                log_map.append('✓ Sorter Map generado.')
            else:
                log_map.append(f'❌ Sorter Map terminó con error (código {code_m})')

        self._json({
            'ok': True,
            'log_gd':    log_gd,
            'log_gantt': log_gantt,
            'log_map':   log_map,
            'warnings':  warnings,
            'has_gantt': has_gantt,
            'has_map':   has_map,
        })

    def _copy_newest(self, pattern, after_mtime):
        """Copy newest matching file from home dir to OUTPUT_DIR. Returns dest path or None."""
        home = os.path.expanduser('~')
        candidates = [f for f in glob.glob(os.path.join(home, pattern))
                      if os.path.getmtime(f) >= after_mtime - 5]
        if not candidates:
            return None
        newest = max(candidates, key=os.path.getmtime)
        dst = os.path.join(OUTPUT_DIR, os.path.basename(newest))
        shutil.copy(newest, dst)
        return dst

    def _respond(self, code, ctype, body):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self._respond(200, 'application/json; charset=utf-8', body)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    PORT = 5001
    server = HTTPServer(('localhost', PORT), Handler)
    print(f'Sorter VDL B2B — Configurador v{APP_VERSION}')
    print(f'→ http://localhost:{PORT}')
    print('Pulsa Ctrl+C para detener.\n')
    threading.Timer(1.2, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nDetenido.')
