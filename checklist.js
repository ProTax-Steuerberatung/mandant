/* ProTax Mandant – Checkliste für den Mandanten als PDF (jsPDF, läuft komplett im Browser) */
(function () {
  'use strict';
  const PETROL = [17, 51, 49], GOLD = [168, 150, 104], GREY = [95, 105, 102], LINE = [185, 198, 195];
  const A4W = 210, A4H = 297, M = 18, CW = A4W - 2 * M, BOTTOM = A4H - 22;

  async function imgData(src) {
    try {
      const r = await fetch(src); if (!r.ok) return null;
      const b = await r.blob();
      return await new Promise(res => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = () => res(null); fr.readAsDataURL(b); });
    } catch (e) { return null; }
  }
  // Zeichen, die die Standardschrift nicht kennt, ersetzen
  const clean = s => String(s ?? '').replace(/[→]/g, '->').replace(/[„“]/g, '"').replace(/[‚‘’]/g, "'");

  async function build({ rec, sections, kanzlei, logo, user }) {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ unit: 'mm', format: 'a4' });
    const logoData = logo ? await imgData(logo) : null;
    let y = M;

    const header = () => {
      if (logoData) { try { pdf.addImage(logoData, 'PNG', M, 12, 34, 16); } catch (e) { } }
      else { pdf.setFont('helvetica', 'bold'); pdf.setFontSize(14); pdf.setTextColor(...PETROL); pdf.text(clean(kanzlei.name || ''), M, 20); }
      pdf.setFont('helvetica', 'normal'); pdf.setFontSize(8); pdf.setTextColor(...GREY);
      [kanzlei.name, kanzlei.adresse, [kanzlei.telefon, kanzlei.email].filter(Boolean).join(' · ')].filter(Boolean)
        .forEach((t, i) => pdf.text(clean(t), A4W - M, 15 + i * 4, { align: 'right' }));
      pdf.setDrawColor(...PETROL); pdf.setLineWidth(0.6); pdf.line(M, 31, A4W - M, 31);
      y = 39;
    };
    const newPage = () => { pdf.addPage(); header(); };
    const need = h => { if (y + h > BOTTOM) newPage(); };

    header();
    pdf.setFont('helvetica', 'bold'); pdf.setFontSize(17); pdf.setTextColor(...PETROL);
    pdf.text('Checkliste für Ihre Mandatsaufnahme', M, y); y += 7;
    pdf.setFont('helvetica', 'normal'); pdf.setFontSize(10); pdf.setTextColor(...GREY);
    pdf.text(clean(`${rec.title} · Stand ${new Date().toLocaleDateString('de-DE')}`), M, y); y += 8;
    pdf.setFontSize(10); pdf.setTextColor(30, 40, 38);
    const intro = 'Bitte ergänzen Sie die folgenden Angaben und legen Sie die genannten Unterlagen bei. Sie können die Checkliste ausdrucken und handschriftlich ausfüllen oder die Angaben einfach in Ihrer Antwort-E-Mail nennen.';
    const il = pdf.splitTextToSize(intro, CW); pdf.text(il, M, y); y += il.length * 4.6 + 4;

    const sectionTitle = t => {
      need(16);
      y += 3;
      pdf.setFillColor(...GOLD); pdf.rect(M, y - 3.2, 2.2, 2.2, 'F');
      pdf.setFont('helvetica', 'bold'); pdf.setFontSize(11.5); pdf.setTextColor(...PETROL);
      pdf.text(clean(t).toUpperCase(), M + 4.5, y - 1); y += 2;
      pdf.setDrawColor(...LINE); pdf.setLineWidth(0.3); pdf.line(M, y, A4W - M, y); y += 6;
    };
    const box = (x, yy, s = 3.6) => { pdf.setDrawColor(...PETROL); pdf.setLineWidth(0.35); pdf.rect(x, yy - s + 0.6, s, s); };
    const lineField = (x, w, label, hint) => {
      pdf.setFont('helvetica', 'normal'); pdf.setFontSize(8); pdf.setTextColor(...GREY);
      pdf.text(clean(label) + (hint ? `  (${hint})` : ''), x, y);
      pdf.setDrawColor(...LINE); pdf.setLineWidth(0.3); pdf.line(x, y + 7.5, x + w, y + 7.5);
    };
    const options = (x, w, label, opts) => {
      pdf.setFont('helvetica', 'normal'); pdf.setFontSize(8); pdf.setTextColor(...GREY);
      pdf.text(clean(label), x, y);
      let ox = x, oy = y + 5.5;
      pdf.setFontSize(9); pdf.setTextColor(30, 40, 38);
      for (const o of opts) {
        const tw = pdf.getTextWidth(clean(o)) + 9;
        if (ox + tw > x + w) { ox = x; oy += 5.5; }
        box(ox, oy); pdf.text(clean(o), ox + 5, oy); ox += tw;
      }
      return oy - y + 4;
    };
    const fieldHeight = f => {
      if (f.t === 'multi' || (f.t === 'select' && f.o.length <= 7)) {
        const opts = f.o; let ox = 0, rows = 1; pdf.setFontSize(9);
        for (const o of opts) { const tw = pdf.getTextWidth(clean(o)) + 9; if (ox + tw > (f.t === 'multi' ? CW : CW / 2 - 4)) { ox = 0; rows++; } ox += tw; }
        return 5.5 * rows + 9;
      }
      return 12;
    };
    const hintFor = f => f.t === 'date' ? 'TT.MM.JJJJ' : f.t === 'stid' ? '11 Ziffern' : '';
    const drawField = (f, x, w) => {
      if (f.t === 'check') { box(x, y + 3); pdf.setFontSize(9); pdf.setTextColor(30, 40, 38); pdf.text(clean(f.l), x + 5, y + 3); return 8; }
      if (f.t === 'multi' || (f.t === 'select' && f.o.length <= 7)) return options(x, w, f.l, f.o.filter(o => o !== 'noch offen' && o !== 'noch zu klären' && o !== 'ohne Anrede'));
      lineField(x, w, f.l, hintFor(f)); return 12;
    };

    for (const sec of sections) {
      sectionTitle(sec.title);
      // einfache Felder zweispaltig, breite Felder und Gruppen über die volle Breite
      let col = 0, rowH = 0;
      const flushRow = () => { if (col) { y += rowH; col = 0; rowH = 0; } };
      for (const it of sec.items) {
        if (it.kind === 'doc') {
          flushRow(); need(7);
          box(M, y); pdf.setFont('helvetica', 'normal'); pdf.setFontSize(10); pdf.setTextColor(30, 40, 38);
          const t = pdf.splitTextToSize(clean(it.doc.l), CW - 8); pdf.text(t, M + 6, y); y += 5.5 * t.length + 1.5;
          continue;
        }
        if (it.kind === 'group') {
          flushRow();
          const g = it.f, n = g.k === 'gf' ? 2 : 3;
          need(10);
          pdf.setFont('helvetica', 'bold'); pdf.setFontSize(10); pdf.setTextColor(...PETROL); pdf.text(clean(g.l), M, y); y += 6;
          for (let i = 0; i < n; i++) {
            const fs = g.fields;
            const h = Math.ceil(fs.length / 2) * 12 + 7;
            need(h);
            pdf.setFont('helvetica', 'normal'); pdf.setFontSize(8.5); pdf.setTextColor(...GOLD); pdf.text(`${({ kinder: 'Kind', gf: 'Geschäftsführer', gs: 'Gesellschafter', beteiligungen: 'Beteiligung' })[g.k] || clean(g.l)} ${i + 1}`, M, y); y += 5;
            fs.forEach((f, j) => {
              const x = M + (j % 2) * (CW / 2 + 4), w = CW / 2 - 4;
              const yy = y; drawField(f, x, w); y = yy;
              if (j % 2 === 1 || j === fs.length - 1) y += 12;
            });
            y += 2;
          }
          continue;
        }
        const f = it.f, full = f.w === 'full' || f.t === 'multi' || f.t === 'textarea';
        const h = fieldHeight(f);
        if (full) { flushRow(); need(h); drawField(f, M, CW); y += h; if (f.t === 'textarea') { pdf.line(M, y + 1.5, A4W - M, y + 1.5); y += 7; } continue; }
        if (col === 0) { need(h); }
        const x = M + col * (CW / 2 + 4), w = CW / 2 - 4;
        drawField(f, x, w); rowH = Math.max(rowH, h);
        col++; if (col === 2) flushRow();
      }
      flushRow();
    }

    need(30); y += 6;
    pdf.setDrawColor(...LINE); pdf.line(M, y, A4W - M, y); y += 6;
    pdf.setFont('helvetica', 'normal'); pdf.setFontSize(9.5); pdf.setTextColor(30, 40, 38);
    const out = pdf.splitTextToSize(clean(`Rücksendung: als Antwort auf unsere E-Mail an ${kanzlei.email || ''} oder persönlich zu Ihrem nächsten Termin. Ihre Ansprechpartnerin bzw. Ihr Ansprechpartner: ${user || ''}.`), CW);
    pdf.text(out, M, y); y += out.length * 4.5 + 10;
    pdf.setDrawColor(...PETROL); pdf.line(M, y, M + 75, y); pdf.line(A4W - M - 75, y, A4W - M, y);
    pdf.setFontSize(8); pdf.setTextColor(...GREY); pdf.text('Ort, Datum', M, y + 4); pdf.text('Unterschrift', A4W - M - 75, y + 4);

    const pages = pdf.getNumberOfPages();
    for (let i = 1; i <= pages; i++) {
      pdf.setPage(i); pdf.setFontSize(7.5); pdf.setTextColor(...GREY);
      pdf.text(clean(`${kanzlei.name || ''} · Vertraulich`), M, A4H - 10);
      pdf.text(`Seite ${i} von ${pages}`, A4W - M, A4H - 10, { align: 'right' });
    }
    return pdf;
  }

  window.CHECKLIST = {
    async blob(opts) { return (await build(opts)).output('blob'); },
    async base64(opts) { const s = (await build(opts)).output('datauristring'); return s.slice(s.indexOf(',') + 1); },
    fileName(rec) { return `Checkliste Mandatsaufnahme ${String(rec.title || '').replace(/[\\/:*?"<>|]/g, ' ').trim()}.pdf`; }
  };
})();
