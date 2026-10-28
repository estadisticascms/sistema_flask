/* Imagen nítida de una tarjeta de reporte para compartir por WhatsApp (usa html2canvas).
   WhatsApp reduce las fotos normales a ~1600 px por el lado más largo: una tarjeta alta (1240 × 3500) llega a ~570 px de
   ancho y se ve borrosa. Por eso "partes" corta la tarjeta en imágenes de máximo 1600 px de alto, siempre entre filas,
   que WhatsApp no necesita reducir. "una" genera la tarjeta completa en alta resolución (para enviar en HD o como documento).
   Siempre se captura al ancho de diseño de la tarjeta, aunque la pantalla (celular) sea más angosta. */
window.ImagenReporte = (function () {
  const MAX_ALTO = 1600;     // px de la imagen; WhatsApp no la reduce si el lado más largo no pasa de esto
  const CORTES = 'tr, .prod, .cab, .kpi2, .mejor, .lider-mes, .armados, .prods .tit, .pie, :scope > *, :scope > * > *';

  async function capturar(el, escala, ancho) {
    let cortes = [];
    const canvas = await html2canvas(el, {
      scale: escala, backgroundColor: '#ffffff', useCORS: true, logging: false,
      windowWidth: Math.max(document.documentElement.clientWidth, ancho + 60),
      onclone: (doc) => {
        const c = doc.getElementById(el.id);
        c.style.width = ancho + 'px'; c.style.maxWidth = 'none'; c.style.margin = '0';
        // Dónde se puede cortar sin partir una fila: el borde inferior de cada bloque
        const top = c.getBoundingClientRect().top;
        cortes = [...new Set([...c.querySelectorAll(CORTES)].map(x => Math.round((x.getBoundingClientRect().bottom - top) * escala)))]
          .filter(y => y > 0).sort((a, b) => a - b);
      },
    });
    return { canvas, cortes };
  }

  function aBlob(canvas) { return new Promise((ok) => canvas.toBlob(ok, 'image/png')); }

  function partir(canvas, cortes, titulo) {
    const piezas = [];
    const objetivo = canvas.height / Math.ceil(canvas.height / (MAX_ALTO - 60));   // partes parejas, sin una última muy corta
    let inicio = 0;
    while (inicio < canvas.height - 2) {
      const banda = piezas.length ? 44 : 0;   // las partes 2, 3… llevan una franja "continuación"
      let fin = Math.min(inicio + objetivo + 150, inicio + MAX_ALTO - banda, canvas.height);
      if (fin < canvas.height) {
        const ok = cortes.filter(y => y > inicio + 200 && y <= fin);
        if (ok.length) fin = ok[ok.length - 1];
      }
      piezas.push([inicio, fin]);
      inicio = fin;
    }
    return piezas.map(([y0, y1], i) => {
      const banda = i ? 44 : 0;
      const c = document.createElement('canvas');
      c.width = canvas.width; c.height = y1 - y0 + banda;
      const g = c.getContext('2d');
      g.fillStyle = '#ffffff'; g.fillRect(0, 0, c.width, c.height);
      if (banda) {
        g.fillStyle = '#027266'; g.fillRect(0, 0, c.width, banda);
        g.fillStyle = '#ffffff'; g.font = 'bold 22px Segoe UI, Arial, sans-serif'; g.textBaseline = 'middle';
        g.fillText(`${titulo} · continuación (${i + 1}/${piezas.length})`, 24, banda / 2);
      }
      g.drawImage(canvas, 0, y0, canvas.width, y1 - y0, 0, banda, canvas.width, y1 - y0);
      return c;
    });
  }

  async function guardar(blobs, nombre) {
    const archivos = blobs.map((b, i) => new File([b], blobs.length > 1 ? nombre.replace(/\.png$/, `_${i + 1}de${blobs.length}.png`) : nombre, { type: 'image/png' }));
    // En el celular: compartir directo a WhatsApp con todas las partes
    if (navigator.canShare && navigator.canShare({ files: archivos })) {
      try { await navigator.share({ files: archivos }); return 'compartido'; }
      catch (e) { if (e.name === 'AbortError') return 'cancelado'; }
    }
    for (const a of archivos) {
      const link = document.createElement('a');
      link.href = URL.createObjectURL(a); link.download = a.name; link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 3000);
      await new Promise(r => setTimeout(r, 400));   // el navegador bloquea descargas seguidas demasiado rápidas
    }
    return 'descargado';
  }

  return {
    /** Tarjeta completa en alta resolución (Blob). */
    async una(el, { ancho = 620, escala = 3 } = {}) {
      return aBlob((await capturar(el, escala, ancho)).canvas);
    },
    /** Partes que WhatsApp no reduce (lista de Blob). */
    async partes(el, { ancho = 620, titulo = '' } = {}) {
      const escala = 2;   // 620 × 2 = 1240 px de ancho: nítido en cualquier celular
      const { canvas, cortes } = await capturar(el, escala, ancho);
      if (canvas.height <= MAX_ALTO) return [await aBlob(canvas)];
      return Promise.all(partir(canvas, cortes, titulo).map(aBlob));
    },
    guardar,
  };
})();
