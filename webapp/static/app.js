const $app = document.getElementById('app');
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const EXT = /\.(jpe?g|png|bmp|tiff?|gif|webp)$/i;
let poll = null;

document.getElementById('logo').onclick = () => { clearInterval(poll); inicio(); };
fetch('/api/estado').then(r => r.json()).then(e => {
  document.getElementById('logo').title = 'Clave: ' + e.clave + ' · Modelo: ' + e.modelo;
  if (e.api_key) { const f = document.createElement('div'); f.className = 'hint'; f.style.textAlign = 'center';
    f.textContent = 'Clave en uso: ' + e.clave + ' · Modelo: ' + e.modelo; document.body.appendChild(f); }
  if (!e.api_key) { const a = document.getElementById('aviso'); a.hidden = false;
    a.textContent = 'Falta configurar ANTHROPIC_API_KEY en el servidor: no se podrán generar informes.'; }
});

function inicio() {
  $app.innerHTML = `<div class="opciones">
    <div class="card opcion" id="nuevo"><h2>➕ Crear nuevo informe</h2><div class="hint">Parte de un archivo plantilla, sus instrucciones y sus fotos</div></div>
    <div class="card opcion" id="exist"><h2>📁 Utilizar existente</h2><div class="hint">Reutiliza un informe ya aprobado con nuevas fotos</div></div></div>`;
  document.getElementById('nuevo').onclick = () => formulario(false);
  document.getElementById('exist').onclick = () => formulario(true);
}

async function formulario(existente) {
  let plantillas = [], sel = null;
  if (existente) {
    plantillas = await (await fetch('/api/plantillas')).json();
    if (!plantillas.length) {
      $app.innerHTML = `<div class="card">Todavía no hay informes aprobados. Cree uno nuevo primero.
        <div class="fila"><button class="pri" id="n">Crear nuevo informe</button></div></div>`;
      document.getElementById('n').onclick = () => formulario(false); return;
    }
  }
  $app.innerHTML = `<div class="card"><h2>${existente ? 'Utilizar informe existente' : 'Crear nuevo informe'}</h2>
    ${existente ? `<label>Informe base</label><div id="lista">${plantillas.map(p => `
      <div class="plantilla" data-id="${p.id}"><div><b>${esc(p.nombre)}</b><br><small>${esc(p.instrucciones.slice(0, 120))}</small></div>
      <button data-del="${p.id}" title="Eliminar">🗑</button></div>`).join('')}</div>` : ''}
    <label>Nombre del ${existente ? 'nuevo ' : 'nuevo '}informe</label><input type="text" id="nombre">
    ${existente ? '' : `<label>Archivo plantilla <span class="hint">(.docx recomendado; también .pdf, .txt, .md)</span></label>
      <input type="file" id="plantilla" accept=".docx,.pdf,.txt,.md">`}
    <label>¿Qué queremos que haga el informe?</label><textarea id="instr"></textarea>
    <label>Carpeta con las fotos <span class="hint">(se incluyen subcarpetas)</span></label>
    <input type="file" id="carpeta" webkitdirectory multiple>
    <div class="hint">o seleccione fotos sueltas:</div><input type="file" id="sueltas" accept="image/*" multiple>
    <div id="cuenta" class="hint"></div><div id="err" class="error"></div>
    <div class="fila"><button class="pri" id="go">Generar informe</button><button id="volver">Cancelar</button></div></div>`;
  document.getElementById('volver').onclick = inicio;
  const instr = document.getElementById('instr');
  if (existente) {
    const marcar = id => { sel = plantillas.find(p => p.id === id);
      document.querySelectorAll('.plantilla').forEach(e => e.classList.toggle('sel', e.dataset.id === id));
      instr.value = sel.instrucciones; };
    document.querySelectorAll('.plantilla').forEach(e => e.onclick = ev => {
      if (ev.target.dataset.del) return; marcar(e.dataset.id); });
    document.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
      if (!confirm('¿Eliminar esta plantilla?')) return;
      await fetch('/api/plantillas/' + b.dataset.del, {method: 'DELETE'}); formulario(true); });
    marcar(plantillas[0].id);
  }
  const fotos = () => [...document.getElementById('carpeta').files, ...document.getElementById('sueltas').files]
    .filter(f => EXT.test(f.name) || f.type.startsWith('image/'));
  const cuenta = () => document.getElementById('cuenta').textContent = fotos().length + ' fotos seleccionadas';
  document.getElementById('carpeta').onchange = cuenta; document.getElementById('sueltas').onchange = cuenta;

  document.getElementById('go').onclick = async () => {
    const err = document.getElementById('err'); err.textContent = '';
    const nombre = document.getElementById('nombre').value.trim(), fs = fotos();
    if (!nombre || !instr.value.trim()) return err.textContent = 'Indique el nombre y lo que debe hacer el informe.';
    if (!existente && !document.getElementById('plantilla').files[0]) return err.textContent = 'Seleccione el archivo plantilla.';
    if (!fs.length) return err.textContent = 'Seleccione la carpeta o las fotos.';
    const fd = new FormData();
    fd.append('nombre', nombre); fd.append('instrucciones', instr.value);
    if (existente) fd.append('plantilla_id', sel.id); else fd.append('plantilla_archivo', document.getElementById('plantilla').files[0]);
    fs.forEach(f => { fd.append('rutas', (f.webkitRelativePath || f.name).split('/').slice(f.webkitRelativePath ? 1 : 0).join('/'));
                      fd.append('fotos', f, f.name); });
    $app.innerHTML = '<div class="card cargando"><span class="spin"></span>Subiendo ' + fs.length + ' fotos…</div>';
    const r = await fetch('/api/trabajos', {method: 'POST', body: fd});
    if (!r.ok) { const d = await r.json().catch(() => ({})); formulario(existente);
      setTimeout(() => document.getElementById('err').textContent = d.detail || 'Error al enviar', 50); return; }
    revision((await r.json()).id);
  };
}

function revision(id, verN) {
  clearInterval(poll);
  const cargar = async () => {
    const j = await (await fetch('/api/trabajos/' + id)).json();
    if (j.estado === 'generando' || j.estado === 'preparando') {
      $app.innerHTML = '<div class="card cargando"><span class="spin"></span>Claude está redactando el informe… (puede tardar un minuto)</div>'; return false; }
    clearInterval(poll); pintar(j, verN); return true;
  };
  cargar().then(fin => { if (!fin) poll = setInterval(cargar, 3000); });
}

function vistaPrevia(id, spec) {
  const foto = f => `/api/trabajos/${id}/fotos/${f.split('/').map(encodeURIComponent).join('/')}`;
  return `<h1>${esc(spec.title)}</h1>` + spec.blocks.map(b => ({
    heading: () => `<h${b.level + 1}>${esc(b.text)}</h${b.level + 1}>`,
    paragraph: () => `<p>${esc(b.text)}</p>`,
    bullets: () => `<ul>${b.items.map(i => `<li>${esc(i)}</li>`).join('')}</ul>`,
    table: () => `<table>${b.header.length ? `<tr>${b.header.map(c => `<th>${esc(c)}</th>`).join('')}</tr>` : ''}${
      b.rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</table>`,
    photo: () => `<figure><img src="${foto(b.file)}" loading="lazy"><figcaption>${esc(b.caption)}</figcaption></figure>`,
    page_break: () => '<hr>',
  }[b.type]())).join('');
}

function pintar(j, verN) {
  if (j.estado === 'error' && !j.versiones.length) {
    $app.innerHTML = `<div class="card"><h2>No se pudo generar el informe</h2><div class="error">${esc(j.error)}</div>
      <div class="fila"><button class="pri" id="re">Reintentar</button><button id="ini">Volver</button></div></div>`;
    document.getElementById('re').onclick = async () => { await fetch(`/api/trabajos/${j.id}/reintentar`, {method: 'POST'}); revision(j.id); };
    document.getElementById('ini').onclick = inicio; return;
  }
  const v = j.versiones.find(x => x.n === verN) || j.versiones[j.versiones.length - 1];
  const ultima = v.n === j.versiones.length;
  $app.innerHTML = `<div class="card"><h2>${esc(j.nombre)}</h2>
    <div class="vers">${j.versiones.map(x => `<button data-v="${x.n}" class="${x.n === v.n ? 'act' : ''}">v${x.n}</button>`).join('')}</div>
    <div class="vista">${vistaPrevia(j.id, v.spec)}</div>
    ${j.error ? `<div class="error">Error en la última modificación: ${esc(j.error)}</div>` : ''}
    ${j.aprobado ? `<p class="hint">✅ Informe aprobado y guardado como plantilla. Ya está disponible en «Utilizar existente».</p>
      <div class="fila"><a href="/api/trabajos/${j.id}/descargar?v=${v.n}"><button class="pri">Descargar Word</button></a><button id="ini">Inicio</button></div>` : `
    <label>¿Es correcto? Si no, indique qué modificar</label>
    <textarea id="cambios" placeholder="Ej.: añade una conclusión por cada fisura; la foto 3 es de un pilar, no de una viga…"></textarea>
    <div id="err" class="error"></div>
    <div class="fila"><button id="mod" ${ultima ? '' : 'disabled'}>Modificar</button>
      <button class="ok" id="apr" ${ultima ? '' : 'disabled'}>✔ Es correcto, aprobar</button>
      <a href="/api/trabajos/${j.id}/descargar?v=${v.n}"><button>Descargar Word</button></a></div>
    ${ultima ? '' : '<div class="hint">Está viendo una versión anterior; seleccione la última para modificar o aprobar.</div>'}`}</div>`;
  document.querySelectorAll('[data-v]').forEach(b => b.onclick = () => pintar(j, +b.dataset.v));
  const ini = document.getElementById('ini'); if (ini) ini.onclick = inicio;
  if (j.aprobado) return;
  document.getElementById('mod').onclick = async () => {
    const c = document.getElementById('cambios').value.trim();
    if (!c) return document.getElementById('err').textContent = 'Escriba qué desea cambiar.';
    const fd = new FormData(); fd.append('cambios', c);
    const r = await fetch(`/api/trabajos/${j.id}/modificar`, {method: 'POST', body: fd});
    if (!r.ok) return document.getElementById('err').textContent = (await r.json()).detail;
    revision(j.id);
  };
  document.getElementById('apr').onclick = async () => {
    const r = await fetch(`/api/trabajos/${j.id}/aprobar`, {method: 'POST'});
    if (!r.ok) return document.getElementById('err').textContent = (await r.json()).detail;
    revision(j.id);
  };
}

inicio();
