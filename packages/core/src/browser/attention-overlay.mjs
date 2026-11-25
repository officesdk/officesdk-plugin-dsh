// Ported from the design handoff for the agent selection box and cursor (v1, 2026-11-04)
// (src/attention-overlay.js), itself an extract of the editor's
// presentation-engine/src/render/overlay/decoration/attention-overlay.ts.
// Only the module wrapper changed (IIFE and globals → ES exports); keep the drawing as it is.
// Updated to the v1.1 handoff (first-generation interaction, 2026-11-05): `focus(…, { selectionProgress })`.
/**
 * Agent attention border — portable extract from presentation.
 *
 * Original:
 *   presentation/packages/presentation-engine/src/render/overlay/decoration/attention-overlay.ts
 *
 * A persistent SVG overlay. Content repaints never rebuild it or reset its
 * clock. Only target changes animate geometry; idle frames move two feathered
 * mask spots around a rounded rectangle. No SVG filters, no stroke-dasharray.
 *
 * Browser: <script src="./attention-overlay.js"></script> then new AttentionOverlay(host).
 * ESM: add `export { AttentionOverlay, ATTENTION_RECIPE };` at the bottom.
 *
 * Usage:
 *   const overlay = new AttentionOverlay(hostElement);
 *   overlay.focus(someElement);          // CSS box of a descendant
 *   overlay.focus({ left, top, width, height }); // CSS px relative to host
 *   overlay.focusWhole();
 *   overlay.hide();
 *   overlay.dispose();
 */
const SVG_NS = "http://www.w3.org/2000/svg";
const GAP_PX = 9;
const RADIUS_PX = 18;
const PERIOD_MS = 3600;
const RESPONSE_SECONDS = 0.36;
const FADE_SECONDS = 0.45;

const COLOR_STOPS = [
  ["0%", "#438bfa"],
  ["32%", "#9464ec"],
  ["64%", "#d96eb9"],
  ["100%", "#18b8c7"],
];

const FADE_STOPS = [
  ["0%", 1],
  ["18%", 0.94],
  ["40%", 0.64],
  ["65%", 0.24],
  ["85%", 0.045],
  ["100%", 0],
];

/** Soft wash + three unmasked strokes. */
const BASE_STROKES = [
  [8, 0.045],
  [4.5, 0.065],
  [1.6, 0.58],
];

/** Masked shimmer stack — concentric strokes, no blur filter. */
const GLOW_STROKES = [
  [14, 0.025],
  [11, 0.05],
  [8, 0.09],
  [5.5, 0.16],
  [3.8, 0.32],
  [2.5, 0.95],
];

let nextOverlayId = 0;

function sameBox(a, b) {
  return a.every((value, index) => value === b[index]);
}

/** A point travelling continuously around a rounded rectangle, in CSS pixels. */
function perimeterPoint(box, phase, radiusPx = RADIUS_PX) {
  const [x, y, width, height] = box;
  const radius = Math.min(radiusPx, width / 2, height / 2);
  const horizontal = width - 2 * radius;
  const vertical = height - 2 * radius;
  const arc = (Math.PI * radius) / 2;
  let distance = phase * (2 * horizontal + 2 * vertical + 4 * arc);
  const line = (length, x1, y1, x2, y2) => {
    if (distance <= length) {
      const ratio = length ? distance / length : 0;
      return [x1 + (x2 - x1) * ratio, y1 + (y2 - y1) * ratio];
    }
    distance -= length;
    return undefined;
  };
  const turn = (cx, cy, start) => {
    if (distance <= arc) {
      const angle = start + (radius ? distance / radius : 0);
      return [cx + radius * Math.cos(angle), cy + radius * Math.sin(angle)];
    }
    distance -= arc;
    return undefined;
  };
  return (
    line(horizontal, x + radius, y, x + width - radius, y) ??
    turn(x + width - radius, y + radius, -Math.PI / 2) ??
    line(vertical, x + width, y + radius, x + width, y + height - radius) ??
    turn(x + width - radius, y + height - radius, 0) ??
    line(horizontal, x + width - radius, y + height, x + radius, y + height) ??
    turn(x + radius, y + height - radius, Math.PI / 2) ??
    line(vertical, x, y + height - radius, x, y + radius) ??
    turn(x + radius, y + radius, Math.PI) ?? [x + radius, y]
  );
}

function create(document, tag, attrs, parent) {
  const element = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) {
    element.setAttribute(key, String(value));
  }
  parent.appendChild(element);
  return element;
}

class AttentionOverlay {
  /**
   * @param {HTMLElement} host  position:relative container. The overlay SVG is
   *   appended as a child and covers the host's padding box.
   */
  constructor(host, { edge = false } = {}) {
    this.edge = edge;
    this.edgeWidth = 1.6;
    if (!host) throw new Error("AttentionOverlay needs a host element");
    this.host = host;
    const document = host.ownerDocument;
    this.window = document.defaultView;
    this.media = this.window?.matchMedia?.("(prefers-reduced-motion: reduce)");
    this.disposed = false;
    this.rects = [];
    this.spots = [];
    this.current = undefined;
    this.target = undefined;
    this.velocity = [0, 0, 0, 0];
    this.clockMs = 0;
    this.lastTime = undefined;
    this.frame = undefined;
    this.fading = false;
    this.opacity = 1;
    this.radiusPx = RADIUS_PX;

    const id = `attention-${++nextOverlayId}`;
    const svg = create(
      document,
      "svg",
      { "data-attention-overlay": "true", "aria-hidden": "true" },
      host,
    );
    Object.assign(svg.style, {
      position: "absolute",
      left: "0",
      top: "0",
      width: "100%",
      height: "100%",
      pointerEvents: "none",
      overflow: edge ? "hidden" : "visible",
      display: "none",
    });
    const defs = create(document, "defs", {}, svg);
    const controls = create(document, "g", {}, svg);

    const colors = create(
      document,
      "linearGradient",
      { id: `${id}-color`, x1: "0%", y1: "0%", x2: "100%", y2: "100%" },
      defs,
    );
    for (const [offset, color] of COLOR_STOPS) {
      create(document, "stop", { offset, "stop-color": color }, colors);
    }

    const fade = create(document, "radialGradient", { id: `${id}-fade` }, defs);
    for (const [offset, opacity] of FADE_STOPS) {
      create(
        document,
        "stop",
        { offset, "stop-color": "white", "stop-opacity": opacity },
        fade,
      );
    }

    this.mask = create(
      document,
      "mask",
      {
        id: `${id}-mask`,
        maskUnits: "userSpaceOnUse",
        x: -30,
        y: -30,
        "mask-type": "alpha",
      },
      defs,
    );
    for (const opacity of [1, 0.55]) {
      this.spots.push(
        create(
          document,
          "circle",
          { cx: 0, cy: 0, r: 100, opacity, fill: `url(#${id}-fade)` },
          this.mask,
        ),
      );
    }

    this.group = create(
      document,
      "g",
      { "data-attention-focus": "true", "pointer-events": "none" },
      controls,
    );

    const rect = (attrs, parent = this.group) => {
      const element = create(
        document,
        "rect",
        { fill: "none", "stroke-linejoin": "round", ...attrs },
        parent,
      );
      this.rects.push(element);
      return element;
    };

    if (!edge) rect({
      fill: `url(#${id}-color)`,
      "fill-opacity": 0.015,
      "data-attention-wash": "true",
    });
    for (const [width, opacity] of (edge ? [[this.edgeWidth, 0.72]] : BASE_STROKES)) {
      rect({
        stroke: `url(#${id}-color)`,
        "stroke-width": width,
        "stroke-opacity": opacity,
        ...(width === 1.6 ? { "data-attention-base": "true" } : {}),
      });
    }

    const glow = create(
      document,
      "g",
      {
        mask: `url(#${id}-mask)`,
        opacity: 0.9,
        "data-attention-shimmer": "true",
      },
      this.group,
    );
    for (const [width, opacity] of (edge ? [[this.edgeWidth, 0.95]] : GLOW_STROKES)) {
      rect(
        {
          stroke: `url(#${id}-color)`,
          "stroke-width": width,
          "stroke-opacity": opacity,
        },
        glow,
      );
    }

    this.svg = svg;
    this.id = id;

    this.resume = () => {
      this.cancelFrame();
      if (this.media?.matches) {
        if (this.fading) this.hide(true);
        else if (this.target) {
          this.current = [...this.target];
          this.velocity = [0, 0, 0, 0];
          this.paintGeometry();
          this.paintLight();
        }
      }
      this.schedule();
    };
    document.addEventListener("visibilitychange", this.resume);
    this.media?.addEventListener?.("change", this.resume);
  }

  /**
   * Outline a descendant element, a CSS-pixel box relative to the host, or
   * the whole host (`{ whole: true }`).
   *
   * @param {HTMLElement | { left: number, top: number, width: number, height: number, whole?: boolean } | undefined} target
   */
  focus(target, { immediate = false, radiusPx = RADIUS_PX, selectionProgress = 1 } = {}) {
    if (this.disposed) return;
    if (!target) {
      this.hide();
      return;
    }
    const hostBox = this.host.getBoundingClientRect();
    const width = hostBox.width;
    const height = hostBox.height;
    if (!(width > 0) || !(height > 0)) {
      this.hide();
      return;
    }

    this.svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    this.mask.setAttribute("width", String(width + 60));
    this.mask.setAttribute("height", String(height + 60));

    const gap = this.edge ? 0 : GAP_PX;
    let box;
    if (target === this.host || target.whole === true) {
      box = [0, 0, width, height];
    } else if (target.nodeType === 1) {
      const r = target.getBoundingClientRect();
      box = [
        r.left - hostBox.left - gap,
        r.top - hostBox.top - gap,
        Math.max(0, r.width) + 2 * gap,
        Math.max(0, r.height) + 2 * gap,
      ];
    } else {
      const left = Number(target.left);
      const top = Number(target.top);
      const w = Number(target.width);
      const h = Number(target.height);
      if (![left, top, w, h].every(Number.isFinite)) {
        this.hide(true);
        return;
      }
      box = [left - gap, top - gap, Math.max(0, w) + 2 * gap, Math.max(0, h) + 2 * gap];
    }

    if (!box.every(Number.isFinite)) {
      this.hide(true);
      return;
    }
    // SVG strokes are centered on the path. Move the path inward by half
    // its width so the visible outer edge is exactly the target boundary.
    if (this.edge) {
      const half = this.edgeWidth / 2;
      box = [box[0] + half, box[1] + half, Math.max(0, box[2] - this.edgeWidth), Math.max(0, box[3] - this.edgeWidth)];
    }
    // Keep the upper-left anchor fixed while the pointer pulls the lower-right corner.
    const reveal = Math.max(0, Math.min(1, selectionProgress));
    box[2] *= reveal;
    box[3] *= reveal;
    const outerRadius = Number.isFinite(radiusPx) ? Math.max(0, radiusPx) : RADIUS_PX;
    const radius = this.edge ? Math.max(0, outerRadius - this.edgeWidth / 2) : outerRadius;
    const radiusChanged = this.radiusPx !== radius;
    this.radiusPx = radius;
    if (this.target && sameBox(box, this.target) && !radiusChanged && !this.fading && (!immediate || sameBox(box, this.current))) return;

    this.target = box;
    if (!this.current || this.media?.matches || immediate || radiusChanged) {
      this.current = [...box];
      this.velocity = [0, 0, 0, 0];
      this.paintGeometry();
    }
    this.fading = false;
    this.opacity = 1;
    this.group.setAttribute("opacity", "1");
    this.svg.style.display = "block";
    this.paintLight();
    this.schedule();
  }

  focusWhole(options) {
    this.focus({ whole: true }, options);
  }

  hide(immediate = false) {
    this.target = undefined;
    if (
      immediate ||
      this.media?.matches ||
      !this.window?.requestAnimationFrame ||
      this.svg.ownerDocument.hidden
    ) {
      this.cancelFrame();
      this.svg.style.display = "none";
      this.current = undefined;
      this.velocity = [0, 0, 0, 0];
      this.fading = false;
    } else if (this.current && !this.fading) {
      this.fading = true;
      this.schedule();
    }
  }

  dispose() {
    this.hide(true);
    this.svg.ownerDocument.removeEventListener("visibilitychange", this.resume);
    this.media?.removeEventListener?.("change", this.resume);
    this.svg.remove();
    this.disposed = true;
  }

  paintGeometry() {
    if (!this.current) return;
    const [x, y, width, height] = this.current;
    const radius = Math.min(this.radiusPx, width / 2, height / 2);
    for (const rect of this.rects) {
      rect.setAttribute("x", String(x));
      rect.setAttribute("y", String(y));
      rect.setAttribute("width", String(width));
      rect.setAttribute("height", String(height));
      rect.setAttribute("rx", String(radius));
    }
    const reach = Math.min(155, Math.max(75, (width + height) * 0.35));
    this.spots[0].setAttribute("r", String(reach));
    this.spots[1].setAttribute("r", String(reach * 0.72));
  }

  paintLight() {
    if (!this.current) return;
    const phase = this.media?.matches
      ? 0.18
      : (this.clockMs % PERIOD_MS) / PERIOD_MS;
    this.spots.forEach((spot, index) => {
      const [x, y] = perimeterPoint(this.current, (phase + index * 0.5) % 1, this.radiusPx);
      spot.setAttribute("transform", `translate(${x} ${y})`);
    });
  }

  schedule() {
    if (
      this.frame !== undefined ||
      this.disposed ||
      !this.current ||
      this.media?.matches ||
      this.svg.ownerDocument.hidden ||
      !this.window?.requestAnimationFrame
    )
      return;
    this.lastTime ??= this.window.performance.now();
    this.frame = this.window.requestAnimationFrame(this.tick);
  }

  tick = (now) => {
    this.frame = undefined;
    const dt = Math.max(0, (now - (this.lastTime ?? now)) / 1000);
    this.lastTime = now;
    this.clockMs += dt * 1000;
    if (
      this.current &&
      this.target &&
      (!sameBox(this.current, this.target) ||
        this.velocity.some((value) => value !== 0))
    ) {
      const before = this.current;
      const omega = 7 / RESPONSE_SECONDS;
      const decay = Math.exp(-omega * dt);
      this.current = before.map((position, index) => {
        const target = this.target[index];
        const delta = position - target;
        const a = this.velocity[index] + omega * delta;
        const next = target + (delta + a * dt) * decay;
        const velocity = (this.velocity[index] - omega * a * dt) * decay;
        const settled =
          Math.abs(next - target) < 0.01 && Math.abs(velocity) < 0.05;
        this.velocity[index] = settled ? 0 : velocity;
        return index >= 2
          ? Math.max(1, settled ? target : next)
          : settled
            ? target
            : next;
      });
      if (!sameBox(before, this.current)) this.paintGeometry();
    }
    if (this.fading) {
      this.opacity = Math.max(0, this.opacity - dt / FADE_SECONDS);
      this.group.setAttribute("opacity", String(this.opacity));
      if (this.opacity === 0) {
        this.hide(true);
        return;
      }
    }
    this.paintLight();
    this.schedule();
  };

  cancelFrame() {
    if (this.frame !== undefined) this.window?.cancelAnimationFrame(this.frame);
    this.frame = undefined;
    this.lastTime = undefined;
  }
}

const ATTENTION_RECIPE = {
  gapPx: GAP_PX,
  radiusPx: RADIUS_PX,
  periodMs: PERIOD_MS,
  responseSeconds: RESPONSE_SECONDS,
  fadeSeconds: FADE_SECONDS,
  colorStops: COLOR_STOPS,
  fadeStops: FADE_STOPS,
  baseStrokes: BASE_STROKES,
  glowStrokes: GLOW_STROKES,
};

export { AttentionOverlay, ATTENTION_RECIPE };
