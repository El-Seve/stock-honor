#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""selftest.py - comprueba que el extractor de la nube (Python) da EXACTAMENTE lo mismo que el
extractor local (Excel + C#) sobre libros sinteticos (sin datos reales).

Los resultados esperados (tests/golden/*.json) los genera tests/make_fixtures.ps1 con el extractor
local. Si algo difiere (una libreria cambia, se toca la logica de un solo lado...), este autotest
falla y la Action NO publica nada.

Uso:  python tools/cloud/selftest.py
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, HERE)

import extract_cloud  # noqa: E402

CASES = [
    # (formato, libro sintetico, esperado sin cobertura, esperado con cobertura)
    ("xlsx", "synthetic_stock.xlsx", "xlsx.json", "xlsx_cov.json"),
    ("xlsb", "synthetic_stock.xlsb", "xlsb.json", "xlsb_cov.json"),
]


def first_diff(a, b):
    n = min(len(a), len(b))
    for i in range(n):
        if a[i] != b[i]:
            lo = max(0, i - 60)
            return i, a[lo:i + 60], b[lo:i + 60]
    return n, a[max(0, n - 60):n + 60], b[max(0, n - 60):n + 60]


def main():
    failed = 0
    for kind, book, g_plain, g_cov in CASES:
        book_path = os.path.join(ROOT, "tests", "fixtures", book)
        for cov, gname in ((False, g_plain), (True, g_cov)):
            label = "%s %s" % (kind, "con cobertura" if cov else "sin cobertura")
            g_path = os.path.join(ROOT, "tests", "golden", gname)
            try:
                with open(g_path, "r", encoding="utf-8") as f:
                    expected = f.read()
                got, stats = extract_cloud.extract_text(book_path, cov)
            except Exception as e:  # noqa: BLE001
                failed += 1
                print("FALLO  %-24s error: %s: %s" % (label, type(e).__name__, e))
                continue
            if got == expected:
                print("OK     %-24s identico al extractor local (%d bytes; %d filas, %d productos, %d PDV)" % (
                    label, len(got), stats["kept"], stats["products"], stats["pdvs"]))
            else:
                failed += 1
                i, ga, ex = first_diff(got, expected)
                print("FALLO  %-24s difiere en el caracter %d (len nube=%d, len local=%d)" % (label, i, len(got), len(expected)))
                print("         nube : %r" % ga)
                print("         local: %r" % ex)
    # el mismo dato en .xlsx y .xlsb debe dar el mismo resultado salvo PERIODO/DIA (numero vs texto)
    a, _ = extract_cloud.extract_text(os.path.join(ROOT, "tests", "fixtures", "synthetic_stock.xlsx"), True)
    b, _ = extract_cloud.extract_text(os.path.join(ROOT, "tests", "fixtures", "synthetic_stock.xlsb"), True)
    strip = lambda t: t[t.index('"products"'):]  # noqa: E731
    if strip(a) == strip(b):
        print("OK     xlsx y xlsb dan los mismos datos")
    else:
        failed += 1
        print("FALLO  xlsx y xlsb NO dan los mismos datos")
    if failed:
        print("\nAUTOTEST FALLIDO (%d)." % failed)
        return 1
    print("\nAUTOTEST OK: el extractor de la nube es identico al local en los libros sinteticos.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
