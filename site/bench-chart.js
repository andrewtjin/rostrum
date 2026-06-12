// Progressive enhancement for the comparison-page benchmark charts.
//
// comparison.html ships two STATIC SVG <img> charts (the no-JS truth, also what
// GitHub's README shows). When this script runs it rebuilds each chart live from
// window.ROSTRUM_BENCH (bench-data.js, numbers only — never document names) and
// swaps it into the figure, adding what a static image can't:
//   · hover/tap a dot -> tooltip with that document's words, cards, file size,
//     the tool's time, and Rostrum's time on the SAME document for contrast
//   · click a legend entry -> toggle that tool's dots on/off
// If JS is unavailable the static <img> (with its CSS-only lightbox zoom) simply
// stays — everything on the page still works.
//
// Geometry, colors, and scales intentionally mirror the generated static SVGs
// (competitors' repo gen-bench-svg.mjs) so the swap is visually seamless.
(function () {
    "use strict";
    var DATA = window.ROSTRUM_BENCH;
    if (!DATA || !Array.isArray(DATA.rows) || !DATA.rows.length) return;

    // row layout: [words, cards, bytes, isSpeechDoc, rostrumS, verbatimS, ddS, jaimeS]
    var ROSTRUM_COL = 4;
    var ENGINES = [
        { col: 4, color: "#1668b8", label: "Rostrum" },
        { col: 5, color: "#c92a2a", label: "Verbatim 6.0.0 (its lab floor)" },
        { col: 6, color: "#e8860c", label: "debate-decoded macro" },
        { col: 7, color: "#7048a8", label: "jaime Zapper macro" },
    ];
    var SVG_NS = "http://www.w3.org/2000/svg";
    // same plot box as the static charts
    var W = 960, H = 540, L = 70, R = 24, T = 46, B = 64;

    function el(name, attrs, text) {
        var n = document.createElementNS(SVG_NS, name);
        for (var k in attrs) n.setAttribute(k, attrs[k]);
        if (text != null) n.textContent = text;
        return n;
    }

    // "1.4 s" / "21 s" / "3 min 5 s" / "2 h 59 min" — friendly, no decimal noise
    function fmtTime(s) {
        if (s < 10) return (Math.round(s * 100) / 100) + " s";
        if (s < 60) return Math.round(s) + " s";
        if (s < 3600) {
            var m = Math.floor(s / 60);
            return m + " min " + Math.round(s - m * 60) + " s";
        }
        var h = Math.floor(s / 3600);
        return h + " h " + Math.round((s - h * 3600) / 60) + " min";
    }
    function fmtBytes(b) {
        if (b >= 1048576) return (b / 1048576).toFixed(1) + " MB";
        return Math.round(b / 1024) + " KB";
    }
    function fmtInt(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ","); }

    // Build one interactive chart inside `figure` plotting row[xCol] (log) vs
    // seconds (log). xCol: 0 = words, 1 = cards.
    function build(figure, xCol, xLabel, title) {
        // flatten to one point per (document, tool) measurement
        var pts = [];
        DATA.rows.forEach(function (row) {
            var x = row[xCol];
            if (!x || x <= 0) return;
            ENGINES.forEach(function (eng, ei) {
                var s = row[eng.col];
                if (s && s > 0) pts.push({ x: x, s: s, ei: ei, row: row });
            });
        });
        if (!pts.length) return;

        // fixed log-log domain over ALL points so legend toggles never rescale
        var lx0 = Infinity, lx1 = -Infinity, ly0 = Infinity, ly1 = -Infinity;
        pts.forEach(function (p) {
            var lx = Math.log10(p.x), ly = Math.log10(p.s);
            if (lx < lx0) lx0 = lx; if (lx > lx1) lx1 = lx;
            if (ly < ly0) ly0 = ly; if (ly > ly1) ly1 = ly;
        });
        var sx = function (v) { return L + ((Math.log10(v) - lx0) / (lx1 - lx0)) * (W - L - R); };
        var sy = function (v) { return H - B - ((Math.log10(v) - ly0) / (ly1 - ly0)) * (H - T - B); };

        var svg = el("svg", {
            viewBox: "0 0 " + W + " " + H,
            role: "img",
            "aria-label": figure.querySelector("img") ? figure.querySelector("img").alt : title,
        });
        svg.setAttribute("class", "bench-svg");
        svg.appendChild(el("rect", { width: W, height: H, fill: "#ffffff" }));
        svg.appendChild(el("text", { x: L, y: 24, class: "bench-title" }, title));
        svg.appendChild(el("text", { x: L, y: 41, class: "bench-sub" },
            "Lower is better · log–log · in-process engine timings (the condition most favorable to Verbatim) · " + DATA.date));

        // gridlines: x at powers of 10, y at human time anchors
        for (var e = Math.ceil(lx0); e <= Math.floor(lx1); e++) {
            var v = Math.pow(10, e);
            svg.appendChild(el("line", { x1: sx(v), y1: T, x2: sx(v), y2: H - B, stroke: "#ececec" }));
            svg.appendChild(el("text", { x: sx(v), y: H - B + 18, class: "bench-tick", "text-anchor": "middle" }, fmtInt(v)));
        }
        [[0.1, "0.1s"], [1, "1s"], [10, "10s"], [60, "1 min"], [600, "10 min"], [3600, "1 hour"], [10800, "3 hours"]]
            .forEach(function (a) {
                if (a[0] < Math.pow(10, ly0) || a[0] > Math.pow(10, ly1)) return;
                svg.appendChild(el("line", { x1: L, y1: sy(a[0]), x2: W - R, y2: sy(a[0]), stroke: "#ececec" }));
                svg.appendChild(el("text", { x: L - 6, y: sy(a[0]) + 4, class: "bench-tick", "text-anchor": "end" }, a[1]));
            });

        // points, grouped per engine so a legend toggle is one style flip;
        // precompute screen coords for the nearest-point hover search
        var groups = ENGINES.map(function () { return el("g", {}); });
        pts.forEach(function (p) {
            p.cx = sx(p.x); p.cy = sy(p.s);
            var eng = ENGINES[p.ei];
            var c = p.row[3]
                ? el("circle", { cx: p.cx.toFixed(1), cy: p.cy.toFixed(1), r: 3, fill: eng.color, "fill-opacity": 0.5 })
                : el("circle", { cx: p.cx.toFixed(1), cy: p.cy.toFixed(1), r: 3, fill: "none", stroke: eng.color, "stroke-opacity": 0.45 });
            groups[p.ei].appendChild(c);
        });
        groups.forEach(function (g) { svg.appendChild(g); });

        // halo ring that jumps to the hovered point
        var halo = el("circle", { r: 6, fill: "none", "stroke-width": 2, visibility: "hidden", "pointer-events": "none" });
        svg.appendChild(halo);

        // clickable in-SVG legend (same placement as the static charts)
        var hidden = ENGINES.map(function () { return false; });
        var ly = T + 16;
        ENGINES.forEach(function (eng, ei) {
            var g = el("g", { class: "bench-leg", tabindex: 0, role: "button", "aria-pressed": "true" });
            g.appendChild(el("rect", { x: L + 4, y: ly - 13, width: 250, height: 19, fill: "#fff", "fill-opacity": 0 }));
            g.appendChild(el("circle", { cx: L + 14, cy: ly - 4, r: 5, fill: eng.color }));
            g.appendChild(el("text", { x: L + 26, y: ly, class: "bench-leg-label" }, eng.label));
            var toggle = function () {
                hidden[ei] = !hidden[ei];
                groups[ei].style.display = hidden[ei] ? "none" : "";
                g.style.opacity = hidden[ei] ? 0.35 : 1;
                g.setAttribute("aria-pressed", String(!hidden[ei]));
                hideTip();
            };
            g.addEventListener("click", toggle);
            g.addEventListener("keydown", function (ev) {
                if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); toggle(); }
            });
            svg.appendChild(g);
            ly += 20;
        });
        svg.appendChild(el("text", { x: L + 14, y: ly + 2, class: "bench-note" },
            "solid = delivered speech docs · hollow = prep/backfiles · click a tool to toggle · hover any dot"));
        svg.appendChild(el("text", { x: L + 14, y: ly + 18, class: "bench-note" },
            "not shown: 6 backfiles Verbatim never finished (>45 min, abandoned) and 1 it crashed Word on"));
        svg.appendChild(el("text", { x: (L + W - R) / 2, y: H - 14, class: "bench-axis", "text-anchor": "middle" }, xLabel));
        var yl = el("text", { x: 20, y: (T + H - B) / 2, class: "bench-axis", "text-anchor": "middle" }, "time to hide");
        yl.setAttribute("transform", "rotate(-90 20 " + (T + H - B) / 2 + ")");
        svg.appendChild(yl);

        // host replaces the static <a><img></a>; figcaption stays
        var host = document.createElement("div");
        host.className = "bench-host";
        host.appendChild(svg);
        var tip = document.createElement("div");
        tip.className = "bench-tip";
        tip.setAttribute("hidden", "");
        host.appendChild(tip);

        function hideTip() { tip.setAttribute("hidden", ""); halo.setAttribute("visibility", "hidden"); }

        // nearest-point hover: pointer position -> viewBox coords -> linear scan
        // (a few thousand points is nothing per frame, and it makes 3px dots
        // hittable without per-circle listeners)
        svg.addEventListener("pointermove", function (ev) {
            var rect = svg.getBoundingClientRect();
            var mx = ((ev.clientX - rect.left) / rect.width) * W;
            var my = ((ev.clientY - rect.top) / rect.height) * H;
            var best = null, bestD = 12 * 12; // 12 viewBox-px grab radius
            for (var i = 0; i < pts.length; i++) {
                var p = pts[i];
                if (hidden[p.ei]) continue;
                var dx = p.cx - mx, dy = p.cy - my;
                var d = dx * dx + dy * dy;
                if (d < bestD) { bestD = d; best = p; }
            }
            if (!best) { hideTip(); return; }
            var eng = ENGINES[best.ei];
            var row = best.row;
            var lines = [
                '<strong style="color:' + eng.color + '">' + eng.label + "</strong> — " + fmtTime(best.s),
                fmtInt(row[0]) + " words · " + fmtInt(row[1]) + " cards · " + fmtBytes(row[2]),
                row[3] ? "delivered speech doc" : "prep / backfile",
            ];
            // the contrast line that sells the chart: Rostrum on the same doc
            if (best.ei !== 0 && row[ROSTRUM_COL]) {
                lines.push("Rostrum on this document: " + fmtTime(row[ROSTRUM_COL]));
            }
            tip.innerHTML = lines.join("<br>");
            tip.removeAttribute("hidden");
            // place near the cursor, flipped when it would overflow the chart
            var tx = ((best.cx / W) * rect.width) + 14;
            var ty = ((best.cy / H) * rect.height) - 10;
            tip.style.left = "0"; tip.style.top = "0"; // reset before measuring
            var tw = tip.offsetWidth, th = tip.offsetHeight;
            if (tx + tw > rect.width) tx = ((best.cx / W) * rect.width) - tw - 14;
            if (ty + th > rect.height) ty = rect.height - th - 4;
            if (ty < 0) ty = 4;
            tip.style.left = tx + "px";
            tip.style.top = ty + "px";
            halo.setAttribute("cx", best.cx);
            halo.setAttribute("cy", best.cy);
            halo.setAttribute("stroke", eng.color);
            halo.setAttribute("visibility", "visible");
        });
        svg.addEventListener("pointerleave", hideTip);

        // swap in: drop the zoom-link thumbnail and its now-unreachable lightbox
        var link = figure.querySelector("a.shot-zoom");
        if (link) {
            var lb = document.getElementById(link.getAttribute("href").slice(1));
            if (lb) lb.parentNode.removeChild(lb);
            figure.replaceChild(host, link);
        } else {
            figure.insertBefore(host, figure.firstChild);
        }
    }

    var wordsFig = document.querySelector('figure[data-bench="words"]');
    var cardsFig = document.querySelector('figure[data-bench="cards"]');
    if (wordsFig) build(wordsFig, 0, "document length, words",
        "Time to hide a document — 807 real tournament docs, four tools");
    if (cardsFig) build(cardsFig, 1, "cards in the document (cite paragraphs)",
        "Time to hide vs number of cards — same corpus");
})();
