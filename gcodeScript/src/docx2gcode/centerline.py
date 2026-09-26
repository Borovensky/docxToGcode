"""Осьова лінія гліфа: растеризація → скелет → полілінії.

Це внутрішня реалізація FontEngine у режимі centerline. Зовнішній інтерфейс
пакета — FontEngine.render(), а не ці функції.
"""

from __future__ import annotations

import math

from .config import Config
from .geom import Point, Polyline, dist
from ._fonttools import BasePen

def rasterize(contours: list[Polyline], ppmm: float, pad: int = 3):
    """Растеризує контури за правилом even-odd. Повертає (сітка, ширина, висота, x0, y0)."""
    xs = [p[0] for c in contours for p in c]
    ys = [p[1] for c in contours for p in c]
    if not xs:
        return None
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    w = int(math.ceil((x1 - x0) * ppmm)) + 2 * pad + 1
    h = int(math.ceil((y1 - y0) * ppmm)) + 2 * pad + 1
    grid = bytearray(w * h)

    # Ребра у піксельних координатах.
    edges: list[tuple[float, float, float, float]] = []
    for c in contours:
        n = len(c)
        closed = dist(c[0], c[-1]) < 1e-9
        pts = c[:-1] if closed else c
        n = len(pts)
        for i in range(n):
            ax, ay = pts[i]
            bx, by = pts[(i + 1) % n]
            edges.append(
                ((ax - x0) * ppmm + pad, (ay - y0) * ppmm + pad,
                 (bx - x0) * ppmm + pad, (by - y0) * ppmm + pad)
            )

    for py in range(h):
        yc = py + 0.5
        crossings: list[float] = []
        for ax, ay, bx, by in edges:
            if (ay <= yc) != (by <= yc):
                t = (yc - ay) / (by - ay)
                crossings.append(ax + t * (bx - ax))
        if not crossings:
            continue
        crossings.sort()
        row = py * w
        for i in range(0, len(crossings) - 1, 2):
            sx = int(math.ceil(crossings[i] - 0.5))
            ex = int(math.floor(crossings[i + 1] - 0.5))
            if ex < sx:
                continue
            sx = max(0, sx)
            ex = min(w - 1, ex)
            for px in range(sx, ex + 1):
                grid[row + px] = 1
    return grid, w, h, x0, y0


_N8 = ((0, -1), (1, -1), (1, 0), (1, 1), (0, 1), (-1, 1), (-1, 0), (-1, -1))


def zhang_suen(grid: bytearray, w: int, h: int) -> bytearray:
    """Потоншення Чжана–Суеня до скелета завтовшки 1 піксель."""
    g = bytearray(grid)
    active = [i for i, v in enumerate(g) if v]
    changed = True
    while changed and active:
        changed = False
        for step in (0, 1):
            to_clear: list[int] = []
            still: list[int] = []
            for idx in active:
                if not g[idx]:
                    continue
                x, y = idx % w, idx // w
                if x <= 0 or y <= 0 or x >= w - 1 or y >= h - 1:
                    still.append(idx)
                    continue
                p = [g[(y + dy) * w + (x + dx)] for dx, dy in _N8]
                b = sum(p)
                if b < 2 or b > 6:
                    still.append(idx)
                    continue
                a = sum(1 for i in range(8) if p[i] == 0 and p[(i + 1) % 8] == 1)
                if a != 1:
                    still.append(idx)
                    continue
                # p = [N, NE, E, SE, S, SW, W, NW]
                n, e, s, wst = p[0], p[2], p[4], p[6]
                if step == 0:
                    ok = (n * e * s == 0) and (e * s * wst == 0)
                else:
                    ok = (n * e * wst == 0) and (n * s * wst == 0)
                if ok:
                    to_clear.append(idx)
                else:
                    still.append(idx)
            for idx in to_clear:
                g[idx] = 0
            if to_clear:
                changed = True
            active = still
    return g


def trace_skeleton(skel: bytearray, w: int, h: int) -> list[list[tuple[int, int]]]:
    """Розбирає скелет на послідовності пікселів (гілки та замкнені петлі)."""
    pixels = {i for i, v in enumerate(skel) if v}
    if not pixels:
        return []

    _cache: dict[int, list[int]] = {}

    def neigh(idx: int) -> list[int]:
        """Сусіди з відкиданням надлишкових діагональних зв'язків.

        У 8-зв'язному скелеті сходинка на кшталт «#» по діагоналі вже з'єднана
        через ортогонального сусіда; без цього фільтра ступінь пікселя
        завищується і шлях розпадається на дволанкові уламки.
        """
        cached = _cache.get(idx)
        if cached is not None:
            return cached
        x, y = idx % w, idx // w
        orth: set[tuple[int, int]] = set()
        out: list[int] = []
        for dx, dy in ((0, -1), (1, 0), (0, 1), (-1, 0)):
            nx, ny = x + dx, y + dy
            if 0 <= nx < w and 0 <= ny < h and (ny * w + nx) in pixels:
                orth.add((dx, dy))
                out.append(ny * w + nx)
        for dx, dy in ((1, -1), (1, 1), (-1, 1), (-1, -1)):
            nx, ny = x + dx, y + dy
            if not (0 <= nx < w and 0 <= ny < h):
                continue
            j = ny * w + nx
            if j not in pixels:
                continue
            if (dx, 0) in orth or (0, dy) in orth:
                continue
            out.append(j)
        _cache[idx] = out
        return out

    degree = {i: len(neigh(i)) for i in pixels}
    nodes = {i for i, d in degree.items() if d != 2}
    visited: set[tuple[int, int]] = set()
    paths: list[list[int]] = []

    def walk(start: int, first: int) -> list[int]:
        path = [start, first]
        visited.add((start, first))
        visited.add((first, start))
        prev, cur = start, first
        while degree.get(cur, 0) == 2:
            nxt = [n for n in neigh(cur) if n != prev]
            if not nxt:
                break
            n = nxt[0]
            if (cur, n) in visited:
                break
            visited.add((cur, n))
            visited.add((n, cur))
            path.append(n)
            prev, cur = cur, n
        return path

    for node in sorted(nodes):
        ns = neigh(node)
        if not ns:
            # Ізольований піксель (крапка, тире) — залишаємо як точковий слід.
            paths.append([node, node])
            continue
        for n in ns:
            if (node, n) not in visited:
                paths.append(walk(node, n))

    # Замкнені петлі без вузлів (наприклад, «о»).
    remaining = pixels - {p for path in paths for p in path}
    while remaining:
        start = min(remaining)
        ns = [n for n in neigh(start) if (start, n) not in visited]
        if not ns:
            remaining.discard(start)
            continue
        path = walk(start, ns[0])
        if path[-1] != start and start in neigh(path[-1]):
            path.append(start)
        paths.append(path)
        remaining -= set(path)

    return [[(p % w, p // w) for p in path] for path in paths if len(path) >= 2]


def join_paths(polys: list[Polyline], angle_limit: float = 135.0) -> list[Polyline]:
    """Зшиває штрихи зі спільними кінцями в довші, щоб перо рідше відривалося.

    З'єднуються лише шляхи, які вже дотикаються, тому малюнок не змінюється —
    змінюється тільки порядок і напрямок обходу.
    """
    paths = [list(p) for p in polys if len(p) >= 2]
    if len(paths) < 2:
        return paths
    limit_cos = math.cos(math.radians(angle_limit))

    def kf(pt: Point) -> tuple[float, float]:
        return (round(pt[0], 3), round(pt[1], 3))

    def outward(path: Polyline, at_end: bool) -> Point:
        a, b = (path[-2], path[-1]) if at_end else (path[1], path[0])
        dx, dy = b[0] - a[0], b[1] - a[1]
        n = math.hypot(dx, dy) or 1.0
        return dx / n, dy / n

    alive = dict(enumerate(paths))
    result: list[Polyline] = []
    while alive:
        cur = alive.pop(next(iter(alive)))
        extended = True
        while extended:
            extended = False
            for at_end in (True, False):
                tip = kf(cur[-1] if at_end else cur[0])
                dvec = outward(cur, at_end)
                best, best_cos, best_rev = None, limit_cos, False
                for j, q in alive.items():
                    for q_end in (False, True):
                        if kf(q[-1] if q_end else q[0]) != tip:
                            continue
                        # Орієнтуємо q так, щоб він починався зі спільної точки,
                        # і беремо напрямок, у якому він іде далі від стику.
                        oriented = q[::-1] if q_end else q
                        qd = outward(oriented, at_end=False)
                        qd = (-qd[0], -qd[1])
                        c = dvec[0] * qd[0] + dvec[1] * qd[1]
                        if c > best_cos:
                            best, best_cos, best_rev = j, c, q_end
                if best is None:
                    continue
                q = alive.pop(best)
                if best_rev:
                    q = q[::-1]
                cur = cur + q[1:] if at_end else q[::-1] + cur[1:]
                extended = True

        result.append(cur)
    return result


def _bbox(c: Polyline) -> tuple[float, float, float, float]:
    xs = [p[0] for p in c]
    ys = [p[1] for p in c]
    return min(xs), min(ys), max(xs), max(ys)


def _split_dots(contours: list[Polyline], limit: float) -> tuple[list[Polyline], list[Polyline]]:
    """Відокремлює дрібні самостійні елементи (крапки, двокрапки, крапка над «і»).

    Потоншення перетворює таку цятку на один піксель і вона зникає, тому її
    контур лишаємо як є. Внутрішні контури («вічко» літери «о») не рахуються
    крапками — вони лежать усередині зовнішнього контуру.
    """
    boxes = [_bbox(c) for c in contours]
    dots: list[Polyline] = []
    rest: list[Polyline] = []
    for i, c in enumerate(contours):
        x0, y0, x1, y1 = boxes[i]
        small = max(x1 - x0, y1 - y0) <= limit
        inside = any(
            j != i and bx0 <= x0 and by0 <= y0 and bx1 >= x1 and by1 >= y1
            for j, (bx0, by0, bx1, by1) in enumerate(boxes)
        )
        (dots if (small and not inside) else rest).append(c)
    return dots, rest


def centerline(contours: list[Polyline], cfg: Config) -> list[Polyline]:
    """Перетворює замкнені контури гліфа на одноштрихову осьову лінію."""
    dots, contours = _split_dots(contours, cfg.dot_size)
    if not contours:
        return dots
    raster = rasterize(contours, cfg.raster_ppmm)
    if raster is None:
        return dots
    grid, w, h, x0, y0 = raster
    skel = zhang_suen(grid, w, h)
    paths = trace_skeleton(skel, w, h)

    scale = 1.0 / cfg.raster_ppmm
    pad = 3
    polys: list[Polyline] = []
    for path in paths:
        poly = [((px - pad) * scale + x0, (py - pad) * scale + y0) for px, py in path]
        polys.append(poly)

    # Прибирання коротких відгалужень (артефакти засічок).
    if cfg.prune > 0:
        pruned: list[Polyline] = []
        endpoints: dict[Point, int] = {}
        for p in polys:
            for e in (p[0], p[-1]):
                key = (round(e[0], 4), round(e[1], 4))
                endpoints[key] = endpoints.get(key, 0) + 1
        for p in polys:
            length = sum(dist(p[i], p[i + 1]) for i in range(len(p) - 1))
            if length >= cfg.prune:
                pruned.append(p)
                continue
            a = (round(p[0][0], 4), round(p[0][1], 4))
            b = (round(p[-1][0], 4), round(p[-1][1], 4))
            free_a = endpoints.get(a, 0) == 1
            free_b = endpoints.get(b, 0) == 1
            # Рівно один вільний кінець → це відгалуження, прибираємо.
            # Обидва вільні → самостійний штрих (крапка, тире), лишаємо.
            if free_a != free_b:
                continue
            pruned.append(p)
        polys = pruned

    return join_paths(polys) + dots


class PolygonPen(BasePen):
    """Перо fontTools, що перетворює контури гліфа на полілінії у мм."""

    def __init__(self, glyph_set, scale: float, step: float, tol: float):
        super().__init__(glyph_set)
        self.scale = scale
        self.step = step
        self.tol = tol
        self.contours: list[Polyline] = []
        self._current: Polyline = []

    def _s(self, pt) -> Point:
        return (pt[0] * self.scale, pt[1] * self.scale)

    def _moveTo(self, pt):
        self._flush()
        self._current = [self._s(pt)]

    def _lineTo(self, pt):
        self._current.append(self._s(pt))

    def _curveToOne(self, p1, p2, p3):
        if not self._current:
            return
        p0 = self._current[-1]
        a, b, c = self._s(p1), self._s(p2), self._s(p3)
        n = self._cubic_segments(p0, a, b, c)
        for i in range(1, n + 1):
            t = i / n
            mt = 1 - t
            x = (mt**3 * p0[0] + 3 * mt**2 * t * a[0] + 3 * mt * t**2 * b[0] + t**3 * c[0])
            y = (mt**3 * p0[1] + 3 * mt**2 * t * a[1] + 3 * mt * t**2 * b[1] + t**3 * c[1])
            self._current.append((x, y))

    def _qCurveToOne(self, p1, p2):
        if not self._current:
            return
        p0 = self._current[-1]
        a, b = self._s(p1), self._s(p2)
        n = self._quad_segments(p0, a, b)
        for i in range(1, n + 1):
            t = i / n
            mt = 1 - t
            x = mt**2 * p0[0] + 2 * mt * t * a[0] + t**2 * b[0]
            y = mt**2 * p0[1] + 2 * mt * t * a[1] + t**2 * b[1]
            self._current.append((x, y))

    def _quad_segments(self, p0: Point, p1: Point, p2: Point) -> int:
        chord = dist(p0, p1) + dist(p1, p2)
        n_len = math.ceil(chord / self.step) if self.step > 0 else 1
        d = math.hypot(p0[0] - 2 * p1[0] + p2[0], p0[1] - 2 * p1[1] + p2[1])
        n_tol = math.ceil(math.sqrt(d / (8 * self.tol))) if self.tol > 0 and d > 0 else 1
        return max(1, min(200, max(n_len, n_tol)))

    def _cubic_segments(self, p0: Point, p1: Point, p2: Point, p3: Point) -> int:
        chord = dist(p0, p1) + dist(p1, p2) + dist(p2, p3)
        n_len = math.ceil(chord / self.step) if self.step > 0 else 1
        d1 = math.hypot(p0[0] - 2 * p1[0] + p2[0], p0[1] - 2 * p1[1] + p2[1])
        d2 = math.hypot(p1[0] - 2 * p2[0] + p3[0], p1[1] - 2 * p2[1] + p3[1])
        d = max(d1, d2)
        n_tol = math.ceil(math.sqrt(3 * d / (4 * self.tol))) if self.tol > 0 and d > 0 else 1
        return max(1, min(200, max(n_len, n_tol)))

    def _closePath(self):
        self._flush(close=True)

    def _endPath(self):
        self._flush()

    def _flush(self, close: bool = False):
        pts = self._current
        self._current = []
        if len(pts) < 2:
            return
        if close and dist(pts[0], pts[-1]) > 1e-9:
            pts.append(pts[0])
        self.contours.append(pts)

    def finish(self) -> list[Polyline]:
        self._flush()
        return self.contours
