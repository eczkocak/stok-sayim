/* Karekodsuz kutuların dosyaları (telefon sayfası ve sayım masası birlikte kullanır).
   Google karekodsuz_ ürün + tür + miad ayı başına toplamı verir:
   {gtin, barkod, ad, tur: 'qr' | 'barkod', ay: 'YYYY-AA' | '', adet}.
   - Barkodlu (EAN-13 vb.): Tria için TXT, her kutu bir satır.
   - QR etiketli: Excel (xlsx); aynı barkodun her miadı ayrı satır, adet toplamı.
   xlsx elle yazılır (sıkıştırmasız zip): dış kütüphane gerekmez. */
(function (w) {
  const sirala = (a, b) => (a.ad || '~' + a.barkod).localeCompare(b.ad || '~' + b.barkod, 'tr') || a.ay.localeCompare(b.ay);

  w.barkodTxt = satirlar => satirlar.filter(s => s.tur === 'barkod').sort(sirala)
    .flatMap(s => Array(s.adet).fill(s.barkod)).join('\r\n');

  w.qrSatirlari = satirlar => satirlar.filter(s => s.tur === 'qr').sort(sirala)
    .map(s => [s.barkod, s.ad || '', s.ay ? s.ay.slice(5, 7) + '/' + s.ay.slice(0, 4) : '', s.adet]);

  // ---------------------------------------------------------------- xlsx --
  const xml = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const sutun = i => String.fromCharCode(65 + i);
  const crcTablo = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = crcTablo[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };

  function zip(dosyalar) {          // [[ad, Uint8Array]] -> Uint8Array (yalnız saklama, sıkıştırma yok)
    const parcalar = [], merkez = []; let konum = 0;
    const d = new Date(), saat = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
      gun = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const [ad, veri] of dosyalar) {
      const a = new TextEncoder().encode(ad), crc = crc32(veri);
      const yerel = new DataView(new ArrayBuffer(30));
      [[0, 0x04034b50, 4], [4, 20, 2], [6, 0, 2], [8, 0, 2], [10, saat, 2], [12, gun, 2], [14, crc, 4], [18, veri.length, 4], [22, veri.length, 4], [26, a.length, 2], [28, 0, 2]]
        .forEach(([o, v, n]) => n === 4 ? yerel.setUint32(o, v, true) : yerel.setUint16(o, v, true));
      const m = new DataView(new ArrayBuffer(46));
      [[0, 0x02014b50, 4], [4, 20, 2], [6, 20, 2], [8, 0, 2], [10, 0, 2], [12, saat, 2], [14, gun, 2], [16, crc, 4], [20, veri.length, 4], [24, veri.length, 4], [28, a.length, 2], [30, 0, 2], [32, 0, 2], [34, 0, 2], [36, 0, 2], [38, 0, 4], [42, konum, 4]]
        .forEach(([o, v, n]) => n === 4 ? m.setUint32(o, v, true) : m.setUint16(o, v, true));
      parcalar.push(new Uint8Array(yerel.buffer), a, veri); merkez.push(new Uint8Array(m.buffer), a);
      konum += 30 + a.length + veri.length;
    }
    const mBoy = merkez.reduce((t, p) => t + p.length, 0), son = new DataView(new ArrayBuffer(22));
    [[0, 0x06054b50, 4], [4, 0, 2], [6, 0, 2], [8, dosyalar.length, 2], [10, dosyalar.length, 2], [12, mBoy, 4], [16, konum, 4], [20, 0, 2]]
      .forEach(([o, v, n]) => n === 4 ? son.setUint32(o, v, true) : son.setUint16(o, v, true));
    const hepsi = [...parcalar, ...merkez, new Uint8Array(son.buffer)], cikti = new Uint8Array(hepsi.reduce((t, p) => t + p.length, 0));
    let i = 0; for (const p of hepsi) { cikti.set(p, i); i += p.length; }
    return cikti;
  }

  /* Tek sayfalık xlsx: ilk satır kalın başlık (dondurulmuş); metinYap: o sütunlar metin
     biçiminde (barkodun baştaki sıfırı kalsın); sayılar sayı olarak yazılır. */
  w.xlsxYap = function (sayfaAdi, basliklar, satirlar, genislikler, metinSutunlar) {
    const hucre = (v, r, i, baslik) => {
      const ref = sutun(i) + r;
      if (baslik) return `<c r="${ref}" t="inlineStr" s="1"><is><t>${xml(v)}</t></is></c>`;
      if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"${metinSutunlar.includes(i) ? ' s="2"' : ''}><is><t xml:space="preserve">${xml(v)}</t></is></c>`;
    };
    const sayfa = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols>' + genislikler.map((g, i) => `<col min="${i + 1}" max="${i + 1}" width="${g}" customWidth="1"/>`).join('') + '</cols><sheetData>' +
      `<row r="1">${basliklar.map((b, i) => hucre(b, 1, i, true)).join('')}</row>` +
      satirlar.map((s, j) => `<row r="${j + 2}">${s.map((v, i) => hucre(v, j + 2, i, false)).join('')}</row>`).join('') +
      '</sheetData></worksheet>';
    const stil = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    const t = s => new TextEncoder().encode(s);
    return zip([
      ['[Content_Types].xml', t('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')],
      ['_rels/.rels', t('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')],
      ['xl/workbook.xml', t('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        `<sheets><sheet name="${xml(sayfaAdi)}" sheetId="1" r:id="rId1"/></sheets></workbook>`)],
      ['xl/_rels/workbook.xml.rels', t('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')],
      ['xl/worksheets/sheet1.xml', t(sayfa)],
      ['xl/styles.xml', t(stil)]
    ]);
  };

  w.qrXlsx = satirlar => w.xlsxYap('QR Etiketli', ['Barkod', 'Ürün Adı', 'Miad (Ay/Yıl)', 'Adet'], w.qrSatirlari(satirlar), [16, 44, 14, 8], [0, 2]);
})(window);
