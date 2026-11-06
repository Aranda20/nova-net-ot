/* ==========================================================================
   xlsx-mini.js - Generador de archivos .xlsx reales, sin librerías externas.
   Soporta: celdas combinadas, colores, bordes, alineación, ajuste de texto,
   anchos de columna, altos de fila e impresión ajustada al ancho de la hoja.
   ========================================================================== */
(function (global) {
    const enc = new TextEncoder();

    // ---------- ZIP (sin compresión) ----------
    const crcTable = (() => {
        const t = new Uint32Array(256);
        for (let n = 0; n < 256; n++) {
            let c = n;
            for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            t[n] = c >>> 0;
        }
        return t;
    })();

    function crc32(buf) {
        let c = 0xFFFFFFFF;
        for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
        return (c ^ 0xFFFFFFFF) >>> 0;
    }

    function concat(arrs) {
        const total = arrs.reduce((a, b) => a + b.length, 0);
        const out = new Uint8Array(total);
        let o = 0;
        arrs.forEach(a => { out.set(a, o); o += a.length; });
        return out;
    }

    function zip(files) {
        const chunks = [], central = [];
        let offset = 0;
        const now = new Date();
        const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
        const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

        files.forEach(f => {
            const name = enc.encode(f.name);
            const crc = crc32(f.data);
            const size = f.data.length;

            const lh = new DataView(new ArrayBuffer(30));
            lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
            lh.setUint16(8, 0, true); lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true);
            lh.setUint32(14, crc, true); lh.setUint32(18, size, true); lh.setUint32(22, size, true);
            lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
            chunks.push(new Uint8Array(lh.buffer), name, f.data);

            const ch = new DataView(new ArrayBuffer(46));
            ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
            ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, dosTime, true);
            ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true); ch.setUint32(20, size, true);
            ch.setUint32(24, size, true); ch.setUint16(28, name.length, true);
            ch.setUint32(42, offset, true);
            central.push(new Uint8Array(ch.buffer), name);

            offset += 30 + name.length + size;
        });

        const centralSize = central.reduce((a, b) => a + b.length, 0);
        const end = new DataView(new ArrayBuffer(22));
        end.setUint32(0, 0x06054b50, true);
        end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
        end.setUint32(12, centralSize, true); end.setUint32(16, offset, true);
        return concat([...chunks, ...central, new Uint8Array(end.buffer)]);
    }

    // ---------- Utilidades XML ----------
    const esc = s => String(s)
        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    function letraCol(n) { // 1 -> A
        let s = '';
        while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
        return s;
    }

    // ---------- Hoja de cálculo ----------
    class LibroXLSX {
        constructor(nombreHoja) {
            this.nombreHoja = (nombreHoja || 'Hoja1').replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31);
            this.celdas = new Map();   // "r,c" -> {v, s}
            this.combinadas = [];
            this.anchos = [];
            this.altos = {};
            this.estilos = [];         // lista de specs (JSON)
            this.mapaEstilos = new Map();
        }

        _estilo(spec) {
            if (!spec) return 0;
            const k = JSON.stringify(spec);
            if (!this.mapaEstilos.has(k)) {
                this.estilos.push(spec);
                this.mapaEstilos.set(k, this.estilos.length); // 0 = estilo por defecto
            }
            return this.mapaEstilos.get(k);
        }

        ancho(col, w) { this.anchos[col] = w; }
        alto(fila, h) { this.altos[fila] = h; }

        // Escribe un valor en el rango (r1,c1)-(r2,c2); combina si el rango tiene más de una celda
        poner(r1, c1, r2, c2, valor, spec) {
            const s = this._estilo(spec);
            for (let r = r1; r <= r2; r++)
                for (let c = c1; c <= c2; c++)
                    this.celdas.set(r + ',' + c, { v: null, s });
            this.celdas.set(r1 + ',' + c1, { v: valor, s });
            if (r2 > r1 || c2 > c1) this.combinadas.push(`${letraCol(c1)}${r1}:${letraCol(c2)}${r2}`);
        }

        _xmlEstilos() {
            const fuentes = [], rellenos = [], bordes = [], xfs = [];
            const idx = (arr, xml) => { let i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
            // Valores obligatorios por defecto
            idx(fuentes, '<font><sz val="10"/><color rgb="FF1E293B"/><name val="Calibri"/><family val="2"/></font>');
            idx(rellenos, '<fill><patternFill patternType="none"/></fill>');
            idx(rellenos, '<fill><patternFill patternType="gray125"/></fill>');
            idx(bordes, '<border><left/><right/><top/><bottom/><diagonal/></border>');

            this.estilos.forEach(sp => {
                const f = idx(fuentes, '<font>' + (sp.b ? '<b/>' : '') + (sp.i ? '<i/>' : '') +
                    `<sz val="${sp.sz || 10}"/><color rgb="FF${sp.color || '1E293B'}"/><name val="Calibri"/><family val="2"/></font>`);
                const r = sp.fill ? idx(rellenos, `<fill><patternFill patternType="solid"><fgColor rgb="FF${sp.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
                const bc = sp.borde || '94A3B8';
                const b = sp.borde === false || sp.borde === undefined ? 0 :
                    idx(bordes, `<border><left style="thin"><color rgb="FF${bc === true ? '94A3B8' : bc}"/></left><right style="thin"><color rgb="FF${bc === true ? '94A3B8' : bc}"/></right><top style="thin"><color rgb="FF${bc === true ? '94A3B8' : bc}"/></top><bottom style="thin"><color rgb="FF${bc === true ? '94A3B8' : bc}"/></bottom><diagonal/></border>`);
                const al = `<alignment horizontal="${sp.h || 'left'}" vertical="${sp.v || 'center'}"${sp.wrap ? ' wrapText="1"' : ''}/>`;
                xfs.push(`<xf numFmtId="0" fontId="${f}" fillId="${r}" borderId="${b}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
            });

            return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                `<fonts count="${fuentes.length}">${fuentes.join('')}</fonts>` +
                `<fills count="${rellenos.length}">${rellenos.join('')}</fills>` +
                `<borders count="${bordes.length}">${bordes.join('')}</borders>` +
                '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
                `<cellXfs count="${xfs.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${xfs.join('')}</cellXfs>` +
                '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
                '</styleSheet>';
        }

        _xmlHoja() {
            let maxR = 1, maxC = 1;
            this.celdas.forEach((_, k) => { const [r, c] = k.split(',').map(Number); if (r > maxR) maxR = r; if (c > maxC) maxC = c; });

            let cols = '';
            for (let c = 1; c <= maxC; c++) if (this.anchos[c]) cols += `<col min="${c}" max="${c}" width="${this.anchos[c]}" customWidth="1"/>`;

            let filas = '';
            for (let r = 1; r <= maxR; r++) {
                let celdasXml = '';
                for (let c = 1; c <= maxC; c++) {
                    const cel = this.celdas.get(r + ',' + c);
                    if (!cel) continue;
                    const ref = letraCol(c) + r, st = cel.s ? ` s="${cel.s}"` : '';
                    if (cel.v === null || cel.v === undefined || cel.v === '') celdasXml += `<c r="${ref}"${st}/>`;
                    else if (typeof cel.v === 'number' && isFinite(cel.v)) celdasXml += `<c r="${ref}"${st}><v>${cel.v}</v></c>`;
                    else celdasXml += `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${esc(cel.v)}</t></is></c>`;
                }
                const ht = this.altos[r] ? ` ht="${this.altos[r]}" customHeight="1"` : '';
                if (celdasXml || ht) filas += `<row r="${r}"${ht}>${celdasXml}</row>`;
            }

            const merge = this.combinadas.length
                ? `<mergeCells count="${this.combinadas.length}">${this.combinadas.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '';

            return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
                `<dimension ref="A1:${letraCol(maxC)}${maxR}"/>` +
                '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>' +
                '<sheetFormatPr defaultRowHeight="15"/>' +
                `<cols>${cols}</cols><sheetData>${filas}</sheetData>${merge}` +
                '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
                '<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="0"/>' +
                '</worksheet>';
        }

        bytes() {
            const x = t => enc.encode(t);
            return zip([
                { name: '[Content_Types].xml', data: x('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>') },
                { name: '_rels/.rels', data: x('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
                { name: 'xl/workbook.xml', data: x(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${esc(this.nombreHoja)}" sheetId="1" r:id="rId1"/></sheets></workbook>`) },
                { name: 'xl/_rels/workbook.xml.rels', data: x('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>') },
                { name: 'xl/styles.xml', data: x(this._xmlEstilos()) },
                { name: 'xl/worksheets/sheet1.xml', data: x(this._xmlHoja()) }
            ]);
        }

        descargar(nombreArchivo) {
            const blob = new Blob([this.bytes()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = nombreArchivo;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 2000);
        }
    }

    global.LibroXLSX = LibroXLSX;
    if (typeof module !== 'undefined') module.exports = { LibroXLSX };
})(typeof window !== 'undefined' ? window : globalThis);
