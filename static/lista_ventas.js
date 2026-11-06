/* Ventana de clientes y productos de un grupo (departamento, zona…) con edición de la clasificación del producto.
   La usan Ventas por departamento (mapa) y Ventas por zona mensual. El servidor devuelve el HTML de la lista
   (templates/ventas_departamento_lista.html).

   ListaVentas.iniciar({
     urlLista,        // ruta que devuelve el HTML: recibe tipo=clientes|productos, ver=<grupo>, cliente=<opcional> y los filtros de base
     base,            // filtros que viajan con cada consulta (desde, hasta, vendedor…)
     urlClasificar,   // POST codigo, segmento, familia, subfamilia
     nombre,          // opcional: cómo mostrar el nombre del grupo en el título
   })
   Botones que abren la ventana: .btn-lista con data-tipo y data-ver (y data-cliente para ir directo a los productos de un cliente). */
window.ListaVentas = (() => {
  const CAMPOS = ['segmento', 'familia', 'subfamilia'];

  function crearVentana() {
    let modal = document.getElementById('modalLista');
    if (modal) return modal;
    modal = document.createElement('div');
    modal.className = 'modal fade';
    modal.id = 'modalLista';
    modal.tabIndex = -1;
    modal.setAttribute('aria-labelledby', 'modalListaTitulo');
    modal.setAttribute('aria-hidden', 'true');
    modal.innerHTML = '<div class="modal-dialog modal-xl modal-dialog-scrollable modal-fullscreen-md-down"><div class="modal-content">' +
      '<div class="modal-header py-2"><h5 class="modal-title" id="modalListaTitulo" style="color: var(--credito)"></h5>' +
      '<button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Cerrar"></button></div>' +
      '<div class="modal-body" id="modalListaCuerpo"></div></div></div>';
    document.body.appendChild(modal);
    return modal;
  }

  function iniciar({urlLista, base, urlClasificar, nombre = (x) => x}) {
    const modal = crearVentana();
    const titulo = modal.querySelector('#modalListaTitulo'), cuerpo = modal.querySelector('#modalListaCuerpo');
    let huboCambios = false;

    async function cargar(tipo, ver, cliente = '') {
      titulo.innerHTML = '<i class="bi ' + (tipo === 'clientes' ? 'bi-people' : 'bi-box-seam') + '" aria-hidden="true"></i> ';
      titulo.append((tipo === 'clientes' ? 'Clientes' : 'Productos') + ' · ' + nombre(ver));
      cuerpo.innerHTML = '<div class="text-center text-muted py-5"><div class="spinner-border spinner-border-sm" role="status"></div> Cargando…</div>';
      bootstrap.Modal.getOrCreateInstance(modal).show();
      try {
        const r = await fetch(urlLista + '?' + new URLSearchParams({...base, tipo, ver, cliente}));
        if (!r.ok) throw new Error(r.status);
        cuerpo.innerHTML = await r.text();
        const buscar = cuerpo.querySelector('[data-buscar-lista]');
        buscar?.addEventListener('input', () => {
          const q = buscar.value.trim().toLowerCase();
          cuerpo.querySelectorAll('tbody tr').forEach(tr => tr.hidden = !!q && !tr.textContent.toLowerCase().includes(q));
        });
        cuerpo.scrollTop = 0;
      } catch (e) { cuerpo.innerHTML = '<div class="alert alert-danger mb-0">No se pudo cargar la lista. Intente de nuevo.</div>'; }
    }

    document.addEventListener('click', (e) => {
      const b = e.target.closest('.btn-lista');
      if (b && !cuerpo.contains(b)) cargar(b.dataset.tipo, b.dataset.ver, b.dataset.cliente || '');
    });
    // Dentro de la ventana: productos de un cliente y regreso a la lista de clientes
    cuerpo.addEventListener('click', (e) => {
      const b = e.target.closest('.btn-productos-cliente, .btn-volver-clientes');
      if (!b) return;
      if (b.classList.contains('btn-productos-cliente')) cargar('productos', b.dataset.ver, b.dataset.cliente);
      else cargar('clientes', b.dataset.ver);
    });

    // ---- Cambiar segmento / familia / subfamilia ----
    const subfamilias = (familia) => {
      const arbol = JSON.parse(cuerpo.querySelector('#arbol-clase')?.textContent || '{}');
      const dl = cuerpo.querySelector('#dl-subfamilia');
      if (dl) dl.replaceChildren(...(arbol[familia] || []).map(s => Object.assign(document.createElement('option'), {value: s})));
    };
    const celdas = (tr) => CAMPOS.map(c => tr.querySelector(`[data-campo="${c}"]`));
    function abrir(tr) {
      celdas(tr).forEach(td => {
        const inp = Object.assign(document.createElement('input'), {type: 'text', value: td.dataset.valor, maxLength: 45,
          className: 'form-control form-control-sm', name: td.dataset.campo});
        inp.setAttribute('list', 'dl-' + td.dataset.campo);
        inp.setAttribute('aria-label', td.dataset.campo.charAt(0).toUpperCase() + td.dataset.campo.slice(1) + ' de ' + tr.dataset.codigo);
        td.replaceChildren(inp);
      });
      const fam = tr.querySelector('input[name=familia]');
      subfamilias(fam.value);
      fam.addEventListener('input', () => subfamilias(fam.value));
      tr.querySelector('.acciones-clase').innerHTML =
        '<div class="d-flex gap-1"><button type="button" class="btn btn-sm btn-verde btn-guardar-clase">Guardar</button>' +
        '<button type="button" class="btn btn-sm btn-outline-secondary btn-cancelar-clase">Cancelar</button></div>';
      tr.classList.add('editando');
      tr.querySelector('input[name=segmento]').focus();
    }
    function cerrar(tr, valores) {
      celdas(tr).forEach(td => {
        const v = valores ? valores[td.dataset.campo] : td.dataset.valor;
        if (valores) td.dataset.valor = v === 'Sin clasificar' ? '' : v;
        td.textContent = v || 'Sin clasificar';
        td.classList.toggle('sin-clase', !td.dataset.valor);
      });
      tr.querySelector('.acciones-clase').innerHTML = '<button type="button" class="btn btn-sm btn-outline-secondary btn-editar-clase" title="Cambiar segmento, familia y subfamilia">' +
        '<i class="bi bi-pencil" aria-hidden="true"></i><span class="visually-hidden">Editar clasificación</span></button>';
      tr.classList.remove('editando');
    }
    async function guardar(tr) {
      const boton = tr.querySelector('.btn-guardar-clase');
      const datos = new FormData();
      datos.append('codigo', tr.dataset.codigo);
      tr.querySelectorAll('input[name]').forEach(i => datos.append(i.name, i.value.trim()));
      boton.disabled = true; boton.textContent = 'Guardando…';
      try {
        const r = await fetch(urlClasificar, {method: 'POST', body: datos, headers: {'X-Requested-With': 'fetch'}});
        const j = await r.json().catch(() => ({ok: false, error: 'Respuesta inesperada del servidor.'}));
        if (!j.ok) throw new Error(j.error || 'No se pudo guardar.');
        cerrar(tr, j);
        tr.classList.add('guardado'); setTimeout(() => tr.classList.remove('guardado'), 2500);
        huboCambios = true;
      } catch (e) {
        alert('No se guardó ' + tr.dataset.codigo + ': ' + e.message);
        boton.disabled = false; boton.textContent = 'Guardar';
      }
    }
    cuerpo.addEventListener('click', (e) => {
      const tr = e.target.closest('tr[data-codigo]');
      if (!tr) return;
      if (e.target.closest('.btn-editar-clase')) abrir(tr);
      else if (e.target.closest('.btn-guardar-clase')) guardar(tr);
      else if (e.target.closest('.btn-cancelar-clase')) cerrar(tr);
    });
    cuerpo.addEventListener('keydown', (e) => {
      const tr = e.target.closest('tr.editando');
      if (!tr || e.target.tagName !== 'INPUT') return;
      if (e.key === 'Enter') { e.preventDefault(); guardar(tr); }
      else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); cerrar(tr); }
    });
    // Si se cambió alguna clasificación, el reporte se recalcula al cerrar la ventana
    modal.addEventListener('hidden.bs.modal', () => { if (huboCambios) location.reload(); });
  }

  return {iniciar};
})();
