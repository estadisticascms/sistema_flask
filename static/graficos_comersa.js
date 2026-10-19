/* Utilidades comunes de Chart.js para los reportes (requiere Chart.js 4 y un contenedor .dash).
   Marcas finas (≤ 24px), extremo redondeado 4px, 2px de separación, rejilla tenue,
   y etiqueta selectiva: solo el total al final de cada barra. */
window.ComersaGraficos = (function () {
  const css = getComputedStyle(document.querySelector('.dash'));
  const v = (n) => css.getPropertyValue(n).trim();
  const C = { credito: v('--credito'), contado: v('--contado'), creditoSuave: v('--credito-suave'),
              contadoSuave: v('--contado-suave'), grid: v('--grid'), ink2: v('--ink-2'), ink3: v('--ink-3'),
              surface: v('--surface') };

  const fmt = (n) => 'C$ ' + new Intl.NumberFormat('es-NI', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);
  const compacto = (n) => 'C$ ' + new Intl.NumberFormat('es-NI', { notation: 'compact', maximumFractionDigits: 1 }).format(n);

  Chart.defaults.font.family = "'Segoe UI', Arial, sans-serif";
  Chart.defaults.font.size = 11;
  Chart.defaults.color = C.ink2;
  Chart.defaults.plugins.legend.display = false;   // la leyenda va en el encabezado de cada tarjeta
  Chart.defaults.plugins.tooltip.callbacks.label = (ctx) =>
    `${ctx.dataset.label}: ${fmt(ctx.parsed[ctx.chart.options.indexAxis === 'y' ? 'x' : 'y'])}`;
  Chart.defaults.maintainAspectRatio = false;

  const ejeValor = () => ({ beginAtZero: true, grid: { color: C.grid, drawTicks: false }, border: { display: false },
                            ticks: { callback: (n) => compacto(n), color: C.ink3, padding: 6 } });
  const ejeCategoria = () => ({ grid: { display: false }, border: { color: C.grid }, ticks: { color: C.ink2 } });

  // Nombres largos en el eje: se recortan (el nombre completo sigue en la tabla y en el tooltip)
  const recortar = (t, max = 24) => (t && t.length > max ? t.slice(0, max - 1) + '…' : t);

  const barra = (label, data, color, extra = {}) => Object.assign({
    label, data, backgroundColor: color, maxBarThickness: 24, borderColor: C.surface, borderWidth: 0,
    borderRadius: 4, borderSkipped: 'start'
  }, extra);

  const totalAlFinal = {
    id: 'totalAlFinal',
    afterDatasetsDraw(chart, _a, opts) {
      if (!opts || !opts.totales) return;
      const { ctx } = chart, horizontal = chart.options.indexAxis === 'y';
      const ultimo = chart.getDatasetMeta(chart.data.datasets.length - 1);
      ctx.save(); ctx.font = "600 11px 'Segoe UI', Arial";
      ultimo.data.forEach((el, i) => {
        const t = opts.totales[i]; if (t == null) return;
        ctx.fillStyle = (opts.colores && opts.colores[i]) || C.ink2;
        const pos = el.tooltipPosition();
        if (horizontal) { ctx.textBaseline = 'middle'; ctx.textAlign = 'left'; ctx.fillText(compacto(t), Math.max(pos.x, el.base) + 6, el.y); }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(compacto(t), el.x, Math.min(pos.y, el.base) - 4); }
      });
      ctx.restore();
    }
  };

  /** Barras horizontales apiladas crédito + contado, con total al final.
   *  resaltar: índice de la barra a destacar (las demás se atenúan). */
  function creditoContado(canvas, nombres, credito, contado, resaltar = null) {
    const atenuar = (fuerte, suave) => nombres.map((_, i) => resaltar === null || i === resaltar ? fuerte : suave);
    return new Chart(canvas, {
      type: 'bar',
      data: { labels: nombres, datasets: [
        barra('Crédito', credito, atenuar(C.credito, C.creditoSuave), { borderRadius: 0 }),
        barra('Contado', contado, atenuar(C.contado, C.contadoSuave), { borderWidth: { left: 2 } }),
      ]},
      options: { indexAxis: 'y', layout: { padding: { right: 56 } },
        scales: { x: Object.assign({ stacked: true }, ejeValor()), y: Object.assign({ stacked: true }, ejeCategoria(),
                  { ticks: { color: C.ink2, font: (ctx) => ({ weight: ctx.index === resaltar ? '700' : '400' }),
                             callback: (_v, i) => recortar(nombres[i]) } }) },
        interaction: { mode: 'index', axis: 'y', intersect: false },
        plugins: { totalAlFinal: { totales: credito.map((c, i) => c + contado[i]),
                                   colores: nombres.map((_, i) => resaltar === null || i === resaltar ? C.ink2 : C.ink3) },
                   tooltip: { callbacks: { footer: (items) => 'Total: ' + fmt(items.reduce((s, it) => s + it.parsed.x, 0)) } } } },
      plugins: [totalAlFinal]
    });
  }

  /** Barras horizontales de una sola serie (marcas, productos). */
  function ranking(canvas, nombres, valores, color = C.credito) {
    return new Chart(canvas, {
      type: 'bar',
      data: { labels: nombres, datasets: [barra('Venta', valores, color)] },
      options: { indexAxis: 'y', layout: { padding: { right: 56 } },
        scales: { x: ejeValor(), y: ejeCategoria() },
        plugins: { totalAlFinal: { totales: valores } } },
      plugins: [totalAlFinal]
    });
  }

  /** Captura un elemento como imagen PNG (para compartir). */
  async function capturar(elemento, claseTemporal) {
    const raiz = document.querySelector('.dash');
    if (claseTemporal) raiz.classList.add(claseTemporal);
    try {
      return await html2canvas(elemento, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false });
    } finally {
      if (claseTemporal) raiz.classList.remove(claseTemporal);
    }
  }

  return { C, fmt, compacto, ejeValor, ejeCategoria, barra, totalAlFinal, creditoContado, ranking, capturar };
})();
