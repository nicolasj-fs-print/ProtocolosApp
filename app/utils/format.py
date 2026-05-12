"""Formateo solo para display (no afecta lógica/match de datos)."""
from __future__ import annotations


def format_serie(value) -> str:
    """`xxxxxxx-xx-xx-x` (guion después de pos 7, 9, 11).

    Robusto a None/NaN y largos cortos: solo inserta los guiones que correspondan.
    Ejemplos:
        '1234567'         -> '1234567'
        '12345678'        -> '1234567-8'
        '1234567890'      -> '1234567-89-0'
        '123456789012'    -> '1234567-89-01-2'
        '1234567890123'   -> '1234567-89-01-23'
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    parts: list[str] = []
    cuts = (7, 9, 11)
    last = 0
    for c in cuts:
        if len(s) > c:
            parts.append(s[last:c])
            last = c
        else:
            break
    parts.append(s[last:])
    return "-".join(p for p in parts if p)
