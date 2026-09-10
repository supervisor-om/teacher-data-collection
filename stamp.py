#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""بصمة الإصدار: تربط data.html بنسخة المصدِّر التي بُنيت معها.

المتصفّح يحتفظ بالملفات عشر دقائق، فقد يستورد مصدِّراً قديماً بينما
القالب يُجلب طازجاً — وفهارس الأنماط بينهما لا تتطابق. البصمة تجعل
الرابط يتغيّر كلّما تغيّر الملف، فيُجلب فوراً بلا انتظار ولا تحديث
يدوي. تُنفَّذ قبل كل نشر يمسّ xlsx-export.js:

    python stamp.py
"""
import hashlib
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent
TARGETS = {'data.html': ['xlsx-export.js']}


def digest_of(mod):
    """بصمة الملف بعد توحيد نهايات الأسطر، فلا تتغيّر لمجرّد اختلاف
       إعداد git بين جهاز وآخر."""
    raw = (ROOT / mod).read_bytes().replace(b'\r\n', b'\n')
    return hashlib.sha256(raw).hexdigest()[:10]


def stamp():
    changed = []
    for page, mods in TARGETS.items():
        path = ROOT / page
        html = before = path.read_text(encoding='utf-8')
        for mod in mods:
            d = digest_of(mod)
            # يُطابَق الرابط ببصمته أو بدونها، فالتشغيل الأول كالتالي
            pat = re.compile(r"(['\"]\./%s)(\?v=[0-9a-f]+)?(['\"])" % re.escape(mod))
            html, n = pat.subn(lambda m: '%s?v=%s%s' % (m.group(1), d, m.group(3)), html)
            if not n:
                sys.exit('لم يُعثر على استيراد %s في %s' % (mod, page))
            print('  %-16s ← %s  (%d موضع)' % (mod, d, n))
        if html != before:
            path.write_text(html, encoding='utf-8')
            changed.append(page)
    print('  حُدّث: ' + (', '.join(changed) if changed else 'لا شيء — البصمة مطابقة'))
    return changed


if __name__ == '__main__':
    stamp()
