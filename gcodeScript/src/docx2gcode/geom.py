"""Геометричні примітиви й згладжування контурів.

Тут лише чиста математика в міліметрах — без DOCX, шрифтів і G-code.
Так само користуються розкладка (ширина тексту), осьова лінія і вивід.
"""

from __future__ import annotations

import math

from .config import TWIPS_PER_MM

Point = tuple[float, float]
Polyline = list[Point]


def mm_from_twips(v: float) -> float:
    return v / TWIPS_PER_MM


def dist(a: Point, b: Point) -> float:
    return math.hypot(b[0] - a[0], b[1] - a[1])


def _lerp(a: Point, b: Point, t: float) -> Point:
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)


def chaikin(pts: Polyline, closed: bool, t: float, corner_cos: float, iters: int) -> Polyline:
    """Згладжування зрізанням кутів (Chaikin) зі збереженням гострих кутів."""
    if t <= 0 or len(pts) < 3:
        return pts
    for _ in range(iters):
        n = len(pts)
        out: Polyline = []
        for i in range(n):
            cur = pts[i]
            has_prev = closed or i > 0
            has_next = closed or i < n - 1
            if not (has_prev and has_next):
                out.append(cur)
                continue
            prev = pts[(i - 1) % n]
            nxt = pts[(i + 1) % n]
            v1 = (cur[0] - prev[0], cur[1] - prev[1])
            v2 = (nxt[0] - cur[0], nxt[1] - cur[1])
            l1 = math.hypot(*v1)
            l2 = math.hypot(*v2)
            if l1 < 1e-12 or l2 < 1e-12:
                out.append(cur)
                continue
            cos_a = (v1[0] * v2[0] + v1[1] * v2[1]) / (l1 * l2)
            if cos_a < corner_cos:  # гострий кут — зберігаємо
                out.append(cur)
                continue
            out.append(_lerp(cur, prev, t))
            out.append(_lerp(cur, nxt, t))
        pts = out
        if len(pts) < 3:
            break
    return pts


def rdp(pts: Polyline, tol: float) -> Polyline:
    """Спрощення Дугласа–Пекера: прибирає точки, що лежать на прямій."""
    if tol <= 0 or len(pts) < 3:
        return pts

    keep = [False] * len(pts)
    keep[0] = keep[-1] = True
    stack = [(0, len(pts) - 1)]
    while stack:
        i0, i1 = stack.pop()
        if i1 <= i0 + 1:
            continue
        ax, ay = pts[i0]
        bx, by = pts[i1]
        dx, dy = bx - ax, by - ay
        norm = math.hypot(dx, dy)
        best_i, best_d = -1, -1.0
        for i in range(i0 + 1, i1):
            px, py = pts[i]
            if norm < 1e-12:
                d = math.hypot(px - ax, py - ay)
            else:
                d = abs(dy * px - dx * py + bx * ay - by * ax) / norm
            if d > best_d:
                best_d, best_i = d, i
        if best_d > tol:
            keep[best_i] = True
            stack.append((i0, best_i))
            stack.append((best_i, i1))
    return [p for p, k in zip(pts, keep) if k]
