// Bloques "atómicos" del dashboard (cada <Card>: fondo blanco, borde y
// esquinas redondeadas) que no deben quedar partidos entre dos páginas del
// PDF. Devuelve sus rangos verticales, en pixeles del canvas ya capturado
// (relativos al propio `container`, escalados por `scale`), ordenados de
// arriba hacia abajo. Se excluyen las tarjetas dentro de algo marcado
// .export-hide (no se capturan, así que no hay nada que proteger ahí).
function getProtectedCardRanges(container: HTMLElement, scale: number) {
  const containerTop = container.getBoundingClientRect().top;
  const cards = Array.from(container.querySelectorAll<HTMLElement>(".rounded-xl.border"));
  return cards
    .filter((el) => !el.closest(".export-hide"))
    .map((el) => {
      const r = el.getBoundingClientRect();
      return {
        top: (r.top - containerTop) * scale,
        bottom: (r.bottom - containerTop) * scale,
      };
    })
    .filter((r) => r.bottom > r.top)
    .sort((a, b) => a.top - b.top);
}

// Captura el elemento del DOM tal cual está renderizado en pantalla (con
// los filtros ya aplicados) y lo vuelca a un PDF paginado en A4. A
// diferencia de un reporte "armado" aparte, esto garantiza que el PDF sea
// un espejo exacto de lo que el usuario está viendo al presionar Export.
//
// Compartido entre el dashboard principal y /agentPerformance (y
// cualquier otra vista futura) para no reimplementar la captura+paginado.
export async function exportElementToPdf(args: { element: HTMLElement; filename: string }) {
  try {
    // html2canvas-pro (fork de html2canvas) entiende los colores oklch/oklab
    // que genera Tailwind v4 en tiempo real; el html2canvas original no los
    // soporta y falla al capturar el dashboard tal cual está en pantalla.
    const [{ default: html2canvas }, { jsPDF }] = await Promise.all([
      import("html2canvas-pro"),
      import("jspdf"),
    ]);

    try {
      // @ts-ignore
      if (document.fonts && document.fonts.ready) {
        // @ts-ignore
        await document.fonts.ready;
      }
    } catch {
      // ignore
    }

    // Importante: NO tocamos el layout del contenedor (nada de
    // padding/margin/resize) antes de capturar. Achicar el contenedor para
    // "hacerle lugar" a la última etiqueta de un eje sonaba bien, pero
    // Recharts recalcula qué ticks mostrar según el ancho disponible, y al
    // achicarlo terminaba OCULTANDO la etiqueta del último punto en vez de
    // sangrarla — dejando puntos de datos sin etiqueta y un hueco raro.
    // En cambio, solo ampliamos el área CAPTURADA (crop) más allá del
    // propio borde del contenedor: cualquier pixel que ya se dibuje un
    // poco más allá de ese borde (ej. la mitad de la última etiqueta de
    // un eje, que Recharts centra sobre el último tick) queda incluido en
    // vez de cortado, sin alterar en nada cómo se ve/mide el dashboard.
    const CAPTURE_SAFETY_MARGIN_PX = 48;
    const SCALE = 2;

    // Los controles de acción (botones, inputs de archivo, selects) no son
    // parte del "informe". Antes se excluían solo de la captura vía
    // `ignoreElements`, pero html2canvas los removía de su clon interno sin
    // colapsar el espacio de la misma forma en que yo lo medía sobre el DOM
    // real — eso corría todo lo de abajo unos pixeles respecto de lo que
    // realmente se capturaba, y desalineaba por completo el cálculo de
    // "dónde termina cada tarjeta". Ahora se ocultan de verdad
    // (display:none) antes de medir y capturar, así ambos ven exactamente
    // el mismo layout colapsado; se restauran apenas termina.
    const hiddenEls = Array.from(args.element.querySelectorAll<HTMLElement>(".export-hide"));
    const previousDisplay = hiddenEls.map((el) => el.style.display);
    hiddenEls.forEach((el) => {
      el.style.display = "none";
    });

    let canvas: HTMLCanvasElement;
    let protectedRanges: ReturnType<typeof getProtectedCardRanges>;
    try {
      canvas = await html2canvas(args.element, {
        scale: SCALE,
        useCORS: true,
        backgroundColor: "#eef2f7",
        width: args.element.scrollWidth + CAPTURE_SAFETY_MARGIN_PX,
      } as any);

      protectedRanges = getProtectedCardRanges(args.element, SCALE);
    } finally {
      hiddenEls.forEach((el, i) => {
        el.style.display = previousDisplay[i];
      });
    }

    const pdf = new jsPDF({ orientation: "p", unit: "mm", format: "a4" });

    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 8;
    const contentW = pageWidth - margin * 2;
    const contentH = pageHeight - margin * 2;

    const pxPerMm = canvas.width / contentW;
    const slicePx = Math.floor(contentH * pxPerMm);

    // 1) Calcular los límites de cada página primero (sin dibujar todavía),
    // para poder ajustar el último antes de renderizar nada.
    const pageSlices: Array<{ sy: number; sh: number }> = [];
    {
      let sy = 0;
      while (sy < canvas.height) {
        let sh = Math.min(slicePx, canvas.height - sy);
        let cutAt = sy + sh;

        // Si el corte "natural" de página cae en medio de una tarjeta, la
        // tarjeta entera se manda a la página siguiente en vez de partirla
        // (el resto de la página actual queda en blanco, preferible a ver
        // un título en una página y el gráfico en la otra). Si la tarjeta
        // ya empezaba antes de esta página (es más alta que una página
        // completa), no hay forma de evitar el corte y se respeta el
        // límite original.
        for (const r of protectedRanges) {
          if (cutAt > r.top && cutAt < r.bottom && r.top > sy) {
            cutAt = Math.min(cutAt, r.top);
          }
        }
        sh = Math.max(1, cutAt - sy);
        pageSlices.push({ sy, sh });
        sy += sh;
      }
    }

    // 2) Si por evitar cortar una tarjeta la última página quedó con un
    // resto mínimo (unos pocos mm de imagen), se sirve mejor sumándolo a
    // la página anterior que dejando una página casi en blanco al final.
    const MIN_TRAILING_PX = Math.round(15 * pxPerMm); // ~15mm de tolerancia
    if (pageSlices.length > 1) {
      const last = pageSlices[pageSlices.length - 1];
      if (last.sh < MIN_TRAILING_PX) {
        pageSlices[pageSlices.length - 2].sh += last.sh;
        pageSlices.pop();
      }
    }

    // 3) Renderizar cada página con los límites ya definitivos.
    pageSlices.forEach(({ sy, sh }, page) => {
      if (page > 0) pdf.addPage();

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = sh;
      const ctx = pageCanvas.getContext("2d");
      if (!ctx) return;
      ctx.drawImage(canvas, 0, sy, canvas.width, sh, 0, 0, canvas.width, sh);

      const pageImg = pageCanvas.toDataURL("image/jpeg", 0.92);
      const imgH = (sh * contentW) / canvas.width;

      pdf.addImage(pageImg, "JPEG", margin, margin, contentW, imgH);
    });

    const blob: Blob = pdf.output("blob");
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = args.filename;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(url), 5000);
  } catch (e: any) {
    console.error("Export PDF failed", e);
    const msg = (e && (e.message || e.toString())) || "Error exportando PDF";
    throw new Error(msg);
  }
}
