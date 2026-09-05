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

    const canvas = await html2canvas(args.element, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#eef2f7",
      width: args.element.scrollWidth + CAPTURE_SAFETY_MARGIN_PX,
      // Los controles de acción (botones, inputs de archivo, selects) no
      // son parte del "informe": se excluyen de la captura via .export-hide.
      ignoreElements: (el: Element) => !!(el as HTMLElement).classList?.contains("export-hide"),
    } as any);

    const pdf = new jsPDF({ orientation: "p", unit: "mm", format: "a4" });

    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 8;
    const contentW = pageWidth - margin * 2;
    const contentH = pageHeight - margin * 2;

    const pxPerMm = canvas.width / contentW;
    const slicePx = Math.floor(contentH * pxPerMm);

    let sy = 0;
    let page = 0;
    while (sy < canvas.height) {
      if (page > 0) pdf.addPage();

      const sh = Math.min(slicePx, canvas.height - sy);
      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = canvas.width;
      pageCanvas.height = sh;
      const ctx = pageCanvas.getContext("2d");
      if (!ctx) break;

      ctx.drawImage(canvas, 0, sy, canvas.width, sh, 0, 0, canvas.width, sh);

      const pageImg = pageCanvas.toDataURL("image/jpeg", 0.92);
      const imgH = (sh * contentW) / canvas.width;

      pdf.addImage(pageImg, "JPEG", margin, margin, contentW, imgH);

      sy += sh;
      page += 1;
    }

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
