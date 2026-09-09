/* بناء ملف Excel بتنسيق ملفَّي المشرف، داخل المتصفّح.
 *
 * القالبان المنشوران (tpl-gov.xlsx و tpl-private.xlsx) مضغوطان بطريقة
 * «التخزين» لا «الانكماش»، فتُقرأ أجزاؤهما كما هي بلا حاجة إلى فكّ ضغط —
 * ولذلك لا تعتمد هذه الوحدة على أي مكتبة خارجية.
 *
 * يُولَّد هذا الملف من layout.json و COLMAP عبر gen_export_js.py، فلا
 * تُحرَّر ثوابته يدوياً.
 */

/* === LAYOUT:BEGIN === */
export const LAYOUT = {
  "حكومية": {
    "file": "tpl-gov.xlsx",
    "headerRow": 6,
    "firstRow": 7,
    "ncols": 25,
    "height": 24.95,
    "styles": {
      "A": "486",
      "B": "487",
      "C": "488",
      "D": "489",
      "E": "490",
      "F": "491",
      "G": "492",
      "H": "493",
      "I": "494",
      "J": "495",
      "K": "496",
      "L": "497",
      "M": "498",
      "N": "499",
      "O": "500",
      "P": "501",
      "Q": "502",
      "R": "503",
      "S": "504",
      "T": "505",
      "U": "506",
      "V": "507",
      "W": "508",
      "X": "509",
      "Y": "510"
    },
    "cols": {
      "A": {
        "k": "الولاية"
      },
      "B": {
        "k": "#م"
      },
      "C": {
        "k": "المشرف"
      },
      "D": null,
      "E": {
        "k": "المدرسة"
      },
      "F": {
        "k": "عدد معلمي الرياضة"
      },
      "G": {
        "k": "رقم الملف"
      },
      "H": {
        "k": "اسم المعلم"
      },
      "I": {
        "k": "عدد الحصص"
      },
      "J": {
        "k": "الفصول التي تدرسها"
      },
      "K": null,
      "L": {
        "k": "الهاتف"
      },
      "M": {
        "k": "السكن الحالي"
      },
      "N": {
        "k": "اسم مدير المدرسة"
      },
      "O": {
        "k": "هاتف مدير المدرسة"
      },
      "P": {
        "k": "سنة التعيين",
        "t": "year"
      },
      "Q": {
        "k": "الدرجة المالية"
      },
      "R": {
        "k": "الرقم المدني"
      },
      "S": {
        "k": "الجامعة"
      },
      "T": {
        "k": "المؤهل"
      },
      "U": {
        "k": "التخصص"
      },
      "V": {
        "k": "الحالة الاجتماعية"
      },
      "W": {
        "k": "تاريخ الميلاد",
        "t": "year"
      },
      "X": {
        "k": "العنوان الدائم"
      },
      "Y": {
        "k": "البريد الالكتروني"
      }
    },
    "numeric": [
      "B",
      "I"
    ]
  },
  "خاصة": {
    "file": "tpl-private.xlsx",
    "headerRow": 3,
    "firstRow": 4,
    "ncols": 30,
    "height": 24.95,
    "styles": {
      "A": "35",
      "B": "63",
      "C": "36",
      "D": "37",
      "E": "38",
      "F": "39",
      "G": "40",
      "H": "41",
      "I": "42",
      "J": "63",
      "K": "43",
      "L": "44",
      "M": "45",
      "N": "46",
      "O": "47",
      "P": "48",
      "Q": "49",
      "R": "50",
      "S": "51",
      "T": "52",
      "U": "53",
      "V": "54",
      "W": "55",
      "X": "56",
      "Y": "57",
      "Z": "58",
      "AA": "59",
      "AB": "60",
      "AC": "61",
      "AD": "62"
    },
    "cols": {
      "A": {
        "k": "الولاية"
      },
      "B": null,
      "C": {
        "k": "المشرف"
      },
      "D": {
        "k": "#م"
      },
      "E": null,
      "F": {
        "k": "المدرسة"
      },
      "G": {
        "k": "الصفوف"
      },
      "H": {
        "k": "عدد معلمي الرياضة"
      },
      "I": {
        "k": "اسم المعلم"
      },
      "J": {
        "k": "مدرجة"
      },
      "K": {
        "k": "عدد الحصص"
      },
      "L": {
        "k": "الصفوف التي يدرسها"
      },
      "M": null,
      "N": {
        "k": "التخصص"
      },
      "O": {
        "k": "الهاتف"
      },
      "P": {
        "k": "الجنسية"
      },
      "Q": {
        "k": "المؤهل"
      },
      "R": {
        "k": "الحالة الاجتماعية"
      },
      "S": {
        "k": "اللغة التي يتم تدريس بها المادة"
      },
      "T": {
        "k": "اسم مدير المدرسة"
      },
      "U": {
        "k": "هاتف مدير المدرسة"
      },
      "V": {
        "k": "جنسية مدير المدرسة"
      },
      "W": {
        "k": "المنهاج"
      },
      "X": {
        "k": "التقويم"
      },
      "Y": {
        "k": "رياضة (11)"
      },
      "Z": {
        "k": "رياضة (12)"
      },
      "AA": {
        "k": "الموافقة على التعيين"
      },
      "AB": {
        "k": "موقع المدرسة"
      },
      "AC": {
        "k": "هاتف آخر للمدرسة"
      },
      "AD": null
    },
    "numeric": [
      "D",
      "H",
      "K"
    ]
  }
};
/* === LAYOUT:END === */

/* -------------------------------------------------------------- CRC32 -- */
const CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

/* ---------------------------------------------------------- قراءة ZIP -- */
/* القالب مخزون بلا ضغط، فيكفي المشي على التواقيع المحلّية. */
function readZip(buf) {
  const b = new Uint8Array(buf);
  const dv = new DataView(buf);
  const dec = new TextDecoder();
  const out = [];
  let i = 0;
  while (i + 4 <= b.length && dv.getUint32(i, true) === 0x04034b50) {
    const method = dv.getUint16(i + 8, true);
    const size = dv.getUint32(i + 18, true);
    const nameLen = dv.getUint16(i + 26, true);
    const extraLen = dv.getUint16(i + 28, true);
    const nameStart = i + 30;
    const dataStart = nameStart + nameLen + extraLen;
    if (method !== 0) throw new Error('القالب مضغوط — يجب أن يكون مخزوناً');
    out.push({ name: dec.decode(b.subarray(nameStart, nameStart + nameLen)),
               data: b.subarray(dataStart, dataStart + size) });
    i = dataStart + size;
  }
  if (!out.length) throw new Error('تعذّرت قراءة القالب');
  return out;
}

/* ---------------------------------------------------------- كتابة ZIP -- */
function writeZip(entries) {
  const enc = new TextEncoder();
  const parts = [];
  const central = [];
  let offset = 0;

  for (const e of entries) {
    const name = enc.encode(e.name);
    const crc = crc32(e.data);
    const lh = new Uint8Array(30 + name.length);
    const lv = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);      /* أدنى إصدار */
    lv.setUint16(6, 0, true);       /* أعلام */
    lv.setUint16(8, 0, true);       /* التخزين */
    lv.setUint16(10, 0, true);      /* الوقت */
    lv.setUint16(12, 0x21, true);   /* التاريخ */
    lv.setUint32(14, crc, true);
    lv.setUint32(18, e.data.length, true);
    lv.setUint32(22, e.data.length, true);
    lv.setUint16(26, name.length, true);
    lv.setUint16(28, 0, true);
    lh.set(name, 30);
    parts.push(lh, e.data);

    const ch = new Uint8Array(46 + name.length);
    const cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0x21, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, e.data.length, true);
    cv.setUint32(24, e.data.length, true);
    cv.setUint16(28, name.length, true);
    cv.setUint32(42, offset, true);
    ch.set(name, 46);
    central.push(ch);

    offset += lh.length + e.data.length;
  }

  const cdSize = central.reduce((n, c) => n + c.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);

  return new Blob([...parts, ...central, eocd],
                  { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ------------------------------------------------------------ الصفوف -- */
const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;' };
const esc = (s) => String(s).replace(/[&<>]/g, (c) => ESC[c]);

function colName(i) {
  let s = '';
  while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = (i - 1 - r) / 26; }
  return s;
}

/* الأصل يكتب عام التعيين بلاحقة «م»، والميلاد يوم/شهر/سنة بلا أصفار بادئة */
const TRANSFORM = {
  year: (v) => (/^\d{4}$/.test(v) ? v + 'م' : v),
  birthdate: (v) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
    return m ? `${+m[3]}/${+m[2]}/${m[1]}` : v;
  }
};

function buildRows(kind, records) {
  const L = LAYOUT[kind];
  const num = new Set(L.numeric);
  const rows = [];
  for (let n = 0; n < records.length; n++) {
    const rec = records[n];
    const rn = L.firstRow + n;
    const cells = [];
    for (let i = 1; i <= L.ncols; i++) {
      const col = colName(i);
      const st = ` s="${L.styles[col] || '0'}"`;
      const ref = `${col}${rn}`;
      const spec = L.cols[col];
      let val;
      if (spec === null || spec === undefined) { cells.push(`<c r="${ref}"${st}/>`); continue; }
      if (spec.k === '#م') val = String(n + 1);
      else {
        val = String(rec[spec.k] == null ? '' : rec[spec.k]).trim();
        if (val && spec.t) val = TRANSFORM[spec.t](val);
      }
      if (!val) cells.push(`<c r="${ref}"${st}/>`);
      else if (num.has(col) && /^\d{1,4}$/.test(val)) cells.push(`<c r="${ref}"${st}><v>${val}</v></c>`);
      else cells.push(`<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${esc(val)}</t></is></c>`);
    }
    rows.push(`<row r="${rn}" spans="1:${L.ncols}" ht="${L.height}" customHeight="1">${cells.join('')}</row>`);
  }
  return rows.join('');
}

function stampLine() {
  const d = new Date();
  return `آخر تحديث: ${d.getDate()}/ ${d.getMonth() + 1} /${d.getFullYear()}م`;
}

/* ------------------------------------------------------------ الواجهة -- */
export async function buildWorkbook(kind, records, baseUrl = '') {
  const L = LAYOUT[kind];
  const res = await fetch(baseUrl + L.file, { cache: 'no-store' });
  if (!res.ok) throw new Error('تعذّر تحميل القالب (' + res.status + ')');
  const entries = readZip(await res.arrayBuffer());

  const dec = new TextDecoder();
  const enc = new TextEncoder();
  const last = Math.max(L.firstRow + records.length - 1, L.headerRow);
  const endCol = colName(L.ncols);

  for (const e of entries) {
    if (e.name === 'xl/worksheets/sheet1.xml') {
      let xml = dec.decode(e.data);
      xml = xml.replace('</sheetData>', buildRows(kind, records) + '</sheetData>');
      xml = xml.replace(/<dimension ref="([A-Z]+)\d+:[A-Z]+\d+"\/>/,
                        (_m, c) => `<dimension ref="${c}1:${endCol}${last}"/>`);
      xml = xml.replace(/<autoFilter ref="([A-Z]+)\d+:[A-Z]+\d+"\/>/,
                        (_m, c) => `<autoFilter ref="${c}${L.headerRow}:${endCol}${last}"/>`);
      xml = xml.replace(/(<t xml:space="preserve">)\s*آخر تحديث[^<]*(<\/t>)/,
                        (_m, a, b) => a + esc(stampLine()) + b);
      e.data = enc.encode(xml);
    } else if (e.name === 'xl/workbook.xml') {
      let xml = dec.decode(e.data);
      xml = xml.replace(/(\$[A-Z]+\$)\d+(:\$[A-Z]+\$)\d+/g,
                        (_m, a, b) => `${a}${L.headerRow}${b}${last}`);
      e.data = enc.encode(xml);
    }
  }
  return writeZip(entries);
}
