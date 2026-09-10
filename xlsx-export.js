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
      "B": "513",
      "C": "488",
      "D": "489",
      "E": "490",
      "F": "514",
      "G": "515",
      "H": "493",
      "I": "516",
      "J": "495",
      "K": "496",
      "L": "524",
      "M": "498",
      "N": "499",
      "O": "525",
      "P": "519",
      "Q": "520",
      "R": "521",
      "S": "504",
      "T": "505",
      "U": "506",
      "V": "507",
      "W": "526",
      "X": "509",
      "Y": "523"
    },
    "sizes": {
      "A": "16",
      "B": "14",
      "C": "16",
      "D": "18",
      "E": "14",
      "F": "14",
      "G": "14",
      "H": "16",
      "I": "14",
      "J": "16",
      "K": "14",
      "L": "16",
      "M": "14",
      "N": "14",
      "O": "16",
      "P": "14",
      "Q": "14",
      "R": "14",
      "S": "14",
      "T": "14",
      "U": "14",
      "V": "16",
      "W": "14",
      "X": "16",
      "Y": "12"
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
        "t": "birthdate"
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
    ],
    "dates": [
      "W"
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
      "D": "64",
      "E": "38",
      "F": "39",
      "G": "40",
      "H": "65",
      "I": "42",
      "J": "63",
      "K": "66",
      "L": "44",
      "M": "45",
      "N": "46",
      "O": "70",
      "P": "48",
      "Q": "49",
      "R": "50",
      "S": "51",
      "T": "52",
      "U": "71",
      "V": "54",
      "W": "55",
      "X": "56",
      "Y": "57",
      "Z": "58",
      "AA": "59",
      "AB": "60",
      "AC": "72",
      "AD": "62"
    },
    "sizes": {
      "A": "16",
      "B": "16",
      "C": "20",
      "D": "14",
      "E": "20",
      "F": "18",
      "G": "16",
      "H": "14",
      "I": "20",
      "J": "16",
      "K": "14",
      "L": "20",
      "M": "18",
      "N": "14",
      "O": "20",
      "P": "16",
      "Q": "16",
      "R": "16",
      "S": "16",
      "T": "16",
      "U": "18",
      "V": "16",
      "W": "16",
      "X": "16",
      "Y": "16",
      "Z": "16",
      "AA": "16",
      "AB": "16",
      "AC": "18",
      "AD": "16"
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
/* تاريخ الميلاد يُحفظ «1976-02-24» من حقل التاريخ في الاستمارة.
   يُكتب رقماً تسلسلياً لا نصّاً كي يراه Excel تاريخاً حقيقياً فيفرزه
   ويرشّحه، ويعرضه بتنسيق العمود dd/mm/yyyy. المبدأ 1899-12-30 هو
   صفر تقويم Excel، والحساب بالتوقيت العالمي فلا يزيحه فارق المنطقة. */
const EPOCH = Date.UTC(1899, 11, 30);
function dateSerial(v) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v).trim());
  if (!m) return null;
  const y = +m[1], mo = +m[2], da = +m[3];
  if (mo < 1 || mo > 12 || da < 1 || da > 31 || y < 1900 || y > 2200) return null;
  const t = Date.UTC(y, mo - 1, da);
  const dt = new Date(t);
  /* يُرفض ما لا يوجد في التقويم مثل 31 فبراير فيبقى نصّاً كما كُتب */
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== da) return null;
  return Math.round((t - EPOCH) / 86400000);
}

const TRANSFORM = {
  year: (v) => (/^\d{4}$/.test(v) ? v + 'م' : v),
  birthdate: (v) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
    return m ? `${+m[3]}/${+m[2]}/${m[1]}` : v;
  }
};

/* نصّ الخليّة مقاطعَ حسب نظام الكتابة: العربية ترث خطّ العمود،
   واللاتينية تُعطى Calibri بحجم العمود نفسه. يُحلّ بذلك المختلط في
   الخليّة الواحدة — «عزان بن قيس الدولية abq» — ولا يُمسّ الخطّ
   العربي، إذ ثبت أنّ تغييره عبر سمة الملف يُسيء إلى العربية. */
const LATIN_RE = /[A-Za-z][A-Za-z0-9 ._@'&/,+-]*[A-Za-z0-9]|[A-Za-z]/g;

function inlineText(val, size) {
  const s = String(val);
  if (!/[A-Za-z]/.test(s)) return `<t xml:space="preserve">${esc(s)}</t>`;

  const runs = [];
  let last = 0;
  for (const m of s.matchAll(LATIN_RE)) {
    if (m.index > last) runs.push([s.slice(last, m.index), false]);
    runs.push([m[0], true]);
    last = m.index + m[0].length;
  }
  if (last < s.length) runs.push([s.slice(last), false]);
  if (runs.length === 1 && !runs[0][1]) return `<t xml:space="preserve">${esc(s)}</t>`;

  const rpr = `<rPr><sz val="${size}"/><color theme="1"/><rFont val="Calibri"/><family val="2"/></rPr>`;
  return runs.map(([t, latin]) =>
    `<r>${latin ? rpr : ''}<t xml:space="preserve">${esc(t)}</t></r>`).join('');
}

function buildRows(kind, records) {
  const L = LAYOUT[kind];
  const num = new Set(L.numeric);
  const dat = new Set(L.dates || []);
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
        if (dat.has(col)) {
          const ser = dateSerial(val);
          if (ser !== null) { cells.push(`<c r="${ref}"${st}><v>${ser}</v></c>`); continue; }
        }
        if (val && spec.t) val = TRANSFORM[spec.t](val);
      }
      if (!val) cells.push(`<c r="${ref}"${st}/>`);
      else if (num.has(col) && /^\d{1,4}$/.test(val)) cells.push(`<c r="${ref}"${st}><v>${val}</v></c>`);
      else cells.push(`<c r="${ref}"${st} t="inlineStr"><is>${inlineText(val, (L.sizes && L.sizes[col]) || 14)}</is></c>`);
    }
    rows.push(`<row r="${rn}" spans="1:${L.ncols}" ht="${L.height}" customHeight="1">${cells.join('')}</row>`);
  }
  return rows.join('');
}

function stampLine() {
  const d = new Date();
  return `آخر تحديث: ${d.getDate()}/ ${d.getMonth() + 1} /${d.getFullYear()}م`;
}

/* القالب يُجلب طازجاً دائماً بينما قد يبقى هذا الملف في ذاكرة
   المتصفّح، فتشير فهارس الأنماط هنا إلى أنماط لم تعد موجودة هناك —
   وXlsx لا يشتكي: ينتج ملفاً يطلب Excel إصلاحه. يُفحص التطابق قبل
   الكتابة ويُصرَّح بالخطأ، خيرٌ من ملفٍ معطوب بيد المشرف. */
function assertStyles(kind, entries, dec) {
  const st = entries.find((e) => e.name === 'xl/styles.xml');
  if (!st) throw new Error('القالب بلا جدول أنماط.');
  const m = /<cellXfs count="(\d+)"/.exec(dec.decode(st.data));
  if (!m) return;
  const have = +m[1];
  const want = Math.max(...Object.values(LAYOUT[kind].styles).map(Number));
  if (want >= have)
    throw new Error('نسخة قديمة محفوظة في المتصفّح. حدّث الصفحة (Ctrl+F5) ثم أعد التصدير.');
}

/* ------------------------------------------------------------ الواجهة -- */
export async function buildWorkbook(kind, records, baseUrl = '') {
  const L = LAYOUT[kind];
  const res = await fetch(baseUrl + L.file, { cache: 'no-store' });
  if (!res.ok) throw new Error('تعذّر تحميل القالب (' + res.status + ')');
  const entries = readZip(await res.arrayBuffer());

  const dec = new TextDecoder();
  assertStyles(kind, entries, dec);

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
