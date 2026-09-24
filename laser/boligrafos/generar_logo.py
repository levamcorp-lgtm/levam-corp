"""Genera el logo LEVAM CORP como formas cerradas (listas para Fill en LightBurn)."""
import math
from shapely.geometry import LineString, box, Polygon
from shapely.ops import unary_union
from shapely import affinity

H, T = 4.0, 0.45          # alto de LEVAM y grosor de trazo (mm)
GAP = 1.9                 # espacio entre letras
h, ts = 1.25, 0.22        # alto de CORP y grosor de trazo
CW, CGAP = 1.15, 0.75     # ancho de letra CORP y espacio


def stroke(pts, t):
    return LineString(pts).buffer(t / 2, cap_style="flat", join_style="mitre", mitre_limit=10)


def clip(g, w, hh):
    return g.intersection(box(-5, 0, w + 5, hh))


def big_L(w=3.2):
    return clip(stroke([(T/2, 0), (T/2, H - T/2), (w, H - T/2)], T), w, H), w


def big_E(w=3.2):
    bars = [box(0, 0, w, T), box(0, H/2 - T/2, w * 0.9, H/2 + T/2), box(0, H - T, w, H)]
    return unary_union(bars), w


def big_V(w=3.8, flip=False):
    a = math.atan((w / 2) / H)
    dx = (T / 2) / math.cos(a)
    e = 2.0
    def ext(p, q):
        (x1, y1), (x2, y2) = p, q
        vx, vy = x2 - x1, y2 - y1
        n = math.hypot(vx, vy)
        return [(x1 - vx / n * e, y1 - vy / n * e), (x2 + vx / n * e, y2 + vy / n * e)]
    left = stroke(ext((dx, 0), (w / 2, H - 0.0)), T)
    right = stroke(ext((w - dx, 0), (w / 2, H - 0.0)), T)
    g = clip(unary_union([left, right]), w, H)
    if flip:
        g = affinity.scale(g, 1, -1, origin=(0, H / 2))
    return g, w


def big_M(w=4.4):
    g = stroke([(T/2, H + 2), (T/2, 0), (w/2, H * 0.62), (w - T/2, 0), (w - T/2, H + 2)], T)
    return clip(g, w, H), w


def ellipse_pts(cx, cy, rx, ry, a0, a1, n=64):
    return [(cx + rx * math.cos(math.radians(a0 + (a1 - a0) * i / n)),
             cy + ry * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]


def s_C():
    pts = ellipse_pts(CW/2, h/2, CW/2 - ts/2, h/2 - ts/2, 45, 315)
    return LineString(pts).buffer(ts/2, cap_style="flat"), CW


def s_O():
    pts = ellipse_pts(CW/2, h/2, CW/2 - ts/2, h/2 - ts/2, 0, 360)
    return Polygon(pts).buffer(ts/2).difference(Polygon(pts).buffer(-ts/2)), CW


def bowl(w):
    bh = h * 0.58
    r = (bh - ts/2 - ts/2) / 2
    xc = w - ts/2 - r
    cy = ts/2 + r
    arc = ellipse_pts(xc, cy, r, r, -90, 90, 32)
    return [(ts/2, ts/2), (xc, ts/2)] + arc[1:-1] + [(xc, bh - ts/2), (ts/2, bh - ts/2)], bh, xc


def s_P(w=CW):
    pts, bh, _ = bowl(w)
    g = unary_union([stroke([(ts/2, -1), (ts/2, h + 1)], ts), stroke(pts, ts)])
    return clip(g, w, h), w


def s_R(w=CW):
    pts, bh, xc = bowl(w)
    leg = stroke([(xc - 0.1, bh - ts/2), (w - ts/2 + 0.3, h + 0.5)], ts)
    g = unary_union([stroke([(ts/2, -1), (ts/2, h + 1)], ts), stroke(pts, ts), leg])
    return clip(g, w, h).intersection(box(0, -1, w, h + 1)), w


def build():
    parts, x = [], 0.0
    for fn in (big_L, big_E, big_V, lambda: big_V(flip=True), big_M):
        g, w = fn()
        parts.append(affinity.translate(g, x, 0))
        x += w + GAP
    total_w = x - GAP

    corp_w = 4 * CW + 3 * CGAP
    cx0 = (total_w - corp_w) / 2
    cy0 = H + 0.9
    x = cx0
    for fn in (s_C, s_O, s_R, s_P):
        g, w = fn()
        parts.append(affinity.translate(g, x, cy0))
        x += w + CGAP
    ly = cy0 + h / 2
    lt = 0.18
    parts.append(box(0, ly - lt/2, cx0 - 0.9, ly + lt/2))
    parts.append(box(cx0 + corp_w + 0.9, ly - lt/2, total_w, ly + lt/2))
    logo = unary_union(parts)
    return logo, total_w, cy0 + h


def to_svg(geom, w, hh, path):
    def ring(c):
        return "M" + " L".join(f"{x:.4f},{y:.4f}" for x, y in c.coords) + " Z"
    d = []
    for p in getattr(geom, "geoms", [geom]):
        d.append(ring(p.exterior))
        d.extend(ring(i) for i in p.interiors)
    svg = (f'<?xml version="1.0" encoding="UTF-8"?>\n'
           f'<svg xmlns="http://www.w3.org/2000/svg" width="{w:.3f}mm" height="{hh:.3f}mm" '
           f'viewBox="0 0 {w:.4f} {hh:.4f}">\n'
           f'  <path fill="#000000" fill-rule="evenodd" stroke="none" d="{" ".join(d)}"/>\n</svg>\n')
    open(path, "w").write(svg)


if __name__ == "__main__":
    import sys, os
    out = os.path.dirname(os.path.abspath(__file__))
    logo, w, hh = build()
    for name, target_h in (("levam-boligrafo-plano.svg", 5.0), ("levam-boligrafo-rotary.svg", 8.0)):
        k = target_h / hh
        g = affinity.scale(logo, k, k, origin=(0, 0))
        to_svg(g, w * k, hh * k, os.path.join(out, name))
        print(f"{name}: {w*k:.1f} x {hh*k:.1f} mm")
