"""Genera el texto del reverso del bolígrafo como formas cerradas (Fill en LightBurn).

Letras geométricas de trazo fino, del mismo estilo que "CORP" en el logo.
"""
import math
import os
from shapely.geometry import LineString, Point, Polygon, box
from shapely.ops import unary_union
from shapely import affinity

TS = 0.13        # grosor de trazo relativo al alto de letra
TRACK = 0.30     # espacio entre letras relativo al alto
SPACE = 0.75     # espacio entre palabras relativo al alto

WIDTHS = {"W": 1.25, "H": 0.8, "O": 1.0, "L": 0.62, "E": 0.62, "S": 0.7, "A": 0.9,
          "D": 0.85, "I": 0.0, "T": 0.8, "R": 0.74, "B": 0.72, "U": 0.8, "N": 0.8,
          "C": 0.92, "M": 1.0, "P": 0.7, "V": 0.9, ".": 0.0, "·": 0.0}


def arc(cx, cy, rx, ry, a0, a1, n=48):
    return [(cx + rx * math.cos(math.radians(a0 + (a1 - a0) * i / n)),
             cy + ry * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]


def extend(pts, e=0.5):
    """Alarga el primer y último segmento para que el recorte deje puntas rectas."""
    def push(p, q):
        vx, vy = p[0] - q[0], p[1] - q[1]
        n = math.hypot(vx, vy) or 1
        return (p[0] + vx / n * e, p[1] + vy / n * e)
    return [push(pts[0], pts[1])] + pts[1:-1] + [push(pts[-1], pts[-2])]


def glyph(ch, h=1.0):
    t = TS * h
    s = t / 2
    w = WIDTHS[ch] * h
    lines, extras = [], []   # polilíneas de eje y formas ya hechas

    if ch == "H":
        lines += [[(s, 0), (s, h)], [(w - s, 0), (w - s, h)], [(s, h / 2), (w - s, h / 2)]]
    elif ch == "I":
        w = t
        lines += [[(s, 0), (s, h)]]
    elif ch == "T":
        lines += [[(0, s), (w, s)], [(w / 2, 0), (w / 2, h)]]
    elif ch == "L":
        lines += [[(s, 0), (s, h - s), (w, h - s)]]
    elif ch == "E":
        lines += [[(w, s), (s, s), (s, h - s), (w, h - s)], [(s, h / 2), (w * 0.88, h / 2)]]
    elif ch == "N":
        lines += [extend([(s, h), (s, 0), (w - s, h), (w - s, 0)])]
    elif ch == "M":
        lines += [extend([(s, h), (s, 0), (w / 2, h * 0.62), (w - s, 0), (w - s, h)])]
    elif ch == "W":
        lines += [extend([(s, 0), (w * 0.26, h), (w / 2, h * 0.3), (w * 0.74, h), (w - s, 0)])]
    elif ch == "V":
        lines += [extend([(s, 0), (w / 2, h), (w - s, 0)])]
    elif ch == "A":
        lines += [extend([(s, h), (w / 2, 0), (w - s, h)])]
        y = h * 0.66
        xl = s + (w / 2 - s) * (1 - y / h)
        lines += [[(xl, y), (w - xl, y)]]
    elif ch == "O":
        ring = Polygon(arc(w / 2, h / 2, w / 2 - s, h / 2 - s, 0, 360, 96))
        extras.append(ring.buffer(s).difference(ring.buffer(-s)))
    elif ch == "C":
        lines += [arc(w / 2, h / 2, w / 2 - s, h / 2 - s, 42, 318, 96)]
    elif ch == "D":
        r = h / 2 - s
        xc = w - s - r
        lines += [[(s, 0), (s, h)], [(s, s), (xc, s)] + arc(xc, h / 2, r, r, -90, 90)[1:-1] + [(xc, h - s), (s, h - s)]]
    elif ch in "PR":
        bh = h * 0.56
        r = (bh - t) / 2
        xc = (w if ch == "P" else w * 0.97) - s - r
        lines += [[(s, 0), (s, h)], [(s, s), (xc, s)] + arc(xc, s + r, r, r, -90, 90)[1:-1] + [(xc, bh - s), (s, bh - s)]]
        if ch == "R":
            lines += [[(xc - 0.05 * h, bh - s), (w - s + 0.2 * h, h + 0.35 * h)]]
    elif ch == "B":
        hm = h * 0.47
        rt = (hm - s - s) / 2
        xt = w * 0.9 - s - rt
        rb = (h - s - hm) / 2
        xb = w - s - rb
        lines += [[(s, 0), (s, h)],
                  [(s, s), (xt, s)] + arc(xt, s + rt, rt, rt, -90, 90)[1:-1] + [(xt, hm), (s, hm)],
                  [(s, hm), (xb, hm)] + arc(xb, hm + rb, rb, rb, -90, 90)[1:-1] + [(xb, h - s), (s, h - s)]]
    elif ch == "U":
        r = w / 2 - s
        cy = h - s - r
        lines += [extend([(s, 0), (s, cy)] + arc(w / 2, cy, r, r, 180, 0)[1:-1] + [(w - s, cy), (w - s, 0)])]
    elif ch == "S":
        rx = w / 2 - s
        ry = h / 4 - s / 2
        top = arc(w / 2, h / 4 + s / 2, rx, ry, -25, -270, 64)
        bot = arc(w / 2, 3 * h / 4 - s / 2, rx, ry, -90, 155, 64)
        lines += [top + bot[1:]]
    elif ch == ".":
        w = t * 1.3
        extras.append(box(0, h - w, w, h))
    elif ch == "·":
        w = t * 2.2
        extras.append(Point(w / 2, h / 2).buffer(w / 2, 32))

    parts = [LineString(p).buffer(s, cap_style="flat", join_style="mitre", mitre_limit=10) for p in lines]
    g = unary_union(parts + extras).intersection(box(-1, 0, w + 1, h))
    return g, w


def text(line, h, track=TRACK):
    parts, x = [], 0.0
    for i, ch in enumerate(line):
        if ch == " ":
            x += SPACE * h - track * h
            continue
        g, w = glyph(ch, h)
        parts.append(affinity.translate(g, x, 0))
        x += w + track * h
    return unary_union(parts), x - track * h


def to_svg(geom, w, hh, path, pad=0.0):
    def ring(c):
        return "M" + " L".join(f"{x + pad:.4f},{y + pad:.4f}" for x, y in c.coords) + " Z"
    d = []
    for p in getattr(geom, "geoms", [geom]):
        d.append(ring(p.exterior))
        d.extend(ring(i) for i in p.interiors)
    W, HH = w + 2 * pad, hh + 2 * pad
    open(path, "w").write(
        f'<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{W:.3f}mm" height="{HH:.3f}mm" viewBox="0 0 {W:.4f} {HH:.4f}">\n'
        f'  <path fill="#000000" fill-rule="evenodd" stroke="none" d="{" ".join(d)}"/>\n</svg>\n')


if __name__ == "__main__":
    out = os.path.dirname(os.path.abspath(__file__))

    # Una línea
    h = 1.8
    g, w = text("WHOLESALE DISTRIBUTION · LEVAMCORP.COM", h, track=0.24)
    to_svg(g, w, h, os.path.join(out, "reverso-una-linea.svg"))
    print(f"reverso-una-linea.svg: {w:.1f} x {h:.1f} mm")

    # Dos líneas, centradas
    h1, h2, gap = 1.8, 1.8, 0.9
    g1, w1 = text("WHOLESALE DISTRIBUTION", h1)
    g2, w2 = text("LEVAMCORP.COM", h2)
    W = max(w1, w2)
    g = unary_union([affinity.translate(g1, (W - w1) / 2, 0),
                     affinity.translate(g2, (W - w2) / 2, h1 + gap)])
    to_svg(g, W, h1 + gap + h2, os.path.join(out, "reverso-dos-lineas.svg"))
    print(f"reverso-dos-lineas.svg: {W:.1f} x {h1 + gap + h2:.1f} mm")
