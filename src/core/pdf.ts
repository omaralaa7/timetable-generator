/** One page of a PDF: a JPEG picture stretched over the whole page. */
export interface PdfPage {
  jpeg: Uint8Array;
  /** Pixel size of the JPEG. */
  width: number;
  height: number;
}

// A4 landscape in PDF points (1/72 inch).
const PAGE_W = 841.89;
const PAGE_H = 595.28;

/**
 * Wrap JPEG pictures into a PDF, one per A4-landscape page. The sheets are drawn on a canvas by
 * the browser (which shapes Arabic correctly) and stored as pictures, so the PDF looks the same
 * everywhere and needs no embedded font; the text is not selectable.
 */
export function imagesToPdf(pages: PdfPage[]): Uint8Array {
  const enc = new TextEncoder();
  const parts: Uint8Array[] = [];
  const offsets: number[] = [];
  let size = 0;
  const push = (chunk: string | Uint8Array) => {
    const bytes = typeof chunk === 'string' ? enc.encode(chunk) : chunk;
    parts.push(bytes);
    size += bytes.length;
  };
  const object = (id: number, body: string, stream?: Uint8Array) => {
    offsets[id] = size;
    push(`${id} 0 obj\n${body}\n`);
    if (stream) {
      push('stream\n');
      push(stream);
      push('\nendstream\n');
    }
    push('endobj\n');
  };

  push('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
  // Objects: 1 catalog, 2 page tree, then per page: page, picture, drawing commands.
  const pageId = (i: number) => 3 + i * 3;
  object(1, '<< /Type /Catalog /Pages 2 0 R >>');
  object(2, `<< /Type /Pages /Kids [${pages.map((_, i) => `${pageId(i)} 0 R`).join(' ')}] /Count ${pages.length} >>`);
  pages.forEach((page, i) => {
    const id = pageId(i);
    const draw = enc.encode(`q ${PAGE_W} 0 0 ${PAGE_H} 0 0 cm /Im0 Do Q`);
    object(id, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Resources << /XObject << /Im0 ${id + 1} 0 R >> >> /Contents ${id + 2} 0 R >>`);
    object(
      id + 1,
      `<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>`,
      page.jpeg,
    );
    object(id + 2, `<< /Length ${draw.length} >>`, draw);
  });
  const count = 3 + pages.length * 3;
  const xref = size;
  push(`xref\n0 ${count}\n0000000000 65535 f \n`);
  for (let id = 1; id < count; id++) push(`${String(offsets[id]).padStart(10, '0')} 00000 n \n`);
  push(`trailer\n<< /Size ${count} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);

  const out = new Uint8Array(size);
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}
