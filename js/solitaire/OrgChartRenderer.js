import * as d3 from 'd3';
import {
    normalizeKey,
    normalizeWs,
    formatMonthYear,
    createOutlookUrl,
    getQueryParam,
    splitNarrativeValues,
    splitValues,
    isMobileDevice,
} from '../shared/utils.js';
import {
    SECOND_LEVEL_LABEL_EXTRA,
    MAX_TEAMS_PER_ROW,
    ROLE_FIELD_WITH_MAPPING,
    COMPANY_FIELD,
    LOCATION_FIELD,
    BUSINESS_FUNCTION_FIELD,
    NEUTRAL_COLOR,
    emailField,
    firstLevelNA,
    secondLevelNA,
} from './constants.js';
import {
    hasTeamDrawerContent,
    getNameFromTitleEl,
    filterOrganizationByStreams,
    getAllowedStreamsSet,
    countTeamsForMemberInOrg,
} from './orgUtils.js';
import { highlightGroup as highlightGroupUtils } from './search.js';
import { BRAND } from '../../brand-specific/brand.js';



const THEME_BOTTOM_PADDING = 40;
const THEME_TOP_PADDING = 110;
const TEAM_TOP_PADDING = 120;
const TEAM_BOTTOM_PADDING_BASE = 28;
const TEAM_MIN_HEIGHT = 220;
const CARD_GAP_Y = 10;
const secondLevelRowPadY = 60;
const ZOOM_MAX_SCALE = 1;
const PAN_MIN_MOVE_PX = 5;
// Zoom threshold below which gesture-LOD activates (FEATURE_LOD must also be true).
// At k < LOD_K_MAX the chart is in overview mode with many cards visible on screen.
const LOD_K_MAX = 0.3;
// Height of a stream box when collapsed — shows header only.
const STREAM_COLLAPSED_HEIGHT = 130;
// Vertical centre of each level's header title (baseline - fontSize × 0.35).
const STREAM_ICON_Y = 75; // stream baseline=105, font=87px
const THEME_ICON_Y  = 62; // theme  baseline=85, font=65px
const TEAM_ICON_Y   = 53; // team   baseline=70, font=50px
// Positions for title/icons when a stream is collapsed to STREAM_COLLAPSED_HEIGHT.
const STREAM_COLLAPSED_ICON_Y  = 65;  // STREAM_COLLAPSED_HEIGHT / 2
const STREAM_COLLAPSED_TITLE_Y = 95;  // 65 + 87 * 0.35 (centres text glyphs in 130px box)

// SVG path data for action icons (centred at 0,0, ~54px bounding box at the chart's SVG scale).
const ICON_CHEVRON_DOWN = 'M -22 -10 L 0 12 L 22 -10';
const ICON_EYE = 'M -27 0 Q -14 -20 0 -20 Q 14 -20 27 0 Q 14 20 0 20 Q -14 20 -27 0 Z M 0 -7.5 A 7.5 7.5 0 1 0 0.01 -7.5';
const ICON_EYE_SLASH = ICON_EYE + ' M -22 -17 L 22 17';
const ICON_INFO = 'M 0 -20 A 20 20 0 1 1 -0.01 -20 M 0 -5 L 0 9';
const ICON_INFO_DOT_CY = -13;

// Generic SVG icon group appended to `parent`, centred at (x, y).
// Pass dotCy to add a filled dot (used by the info icon's "i" dot).
function _appendIconG(parent, cls, tooltip, x, y, pathD, dotCy = null) {
    const g = parent.append('g')
        .attr('class', cls)
        .attr('transform', `translate(${x}, ${y})`)
        .attr('data-tooltip', tooltip)
        .attr('aria-label', tooltip)
        .attr('data-icon-x', x)
        .style('cursor', 'pointer');
    g.append('circle').attr('r', 28).attr('fill', 'transparent');
    g.append('path')
        .attr('d', pathD)
        .attr('fill', 'none')
        .attr('stroke', '#333')
        .attr('stroke-width', '3.5')
        .attr('stroke-linecap', 'round')
        .attr('stroke-linejoin', 'round');
    if (dotCy !== null) {
        g.append('circle').attr('cx', 0).attr('cy', dotCy).attr('r', 3.5).attr('fill', '#333');
    }
    return g;
}

function _contentSelector(streamKey) {
    const base = streamKey.replace(/^stream::/, '');
    return `g[data-key^="theme::${base}::"], g[data-key^="team::${base}::"], g[data-key^="card::${base}::"]`;
}

function getTalentKeringUrl(member, emailFieldName = 'Email') {
    const email = (member?.[emailFieldName] ?? '').toString().trim();
    return `https://kering.eightfold.ai/careerhub/search/people?query=${encodeURIComponent(email)}`;
}

export const byBoost = (boostMap, prefix = '') => ([a], [b]) => {
    const ba = boostMap[prefix + a] ?? null;
    const bb = boostMap[prefix + b] ?? null;
    if (ba !== null && bb !== null) return bb - ba;
    if (ba !== null) return -1;
    if (bb !== null) return 1;
    return a.localeCompare(b, 'en', { sensitivity: 'base' });
};

function communityTooltip(group) {
    return `People sharing a part-time focus on an interest or short-term mission, rather than a full-time daily ${group}.`;
}

function _dominoAccentIconSrc() {
    const t = document.documentElement?.getAttribute('data-theme');
    const dark = t === 'dark' || (!t && window.matchMedia?.('(prefers-color-scheme: dark)')?.matches);
    return dark ? 'assets/domino-accent.svg' : 'assets/domino-accent-light.svg';
}

export class OrgChartRenderer {
    constructor(app) {
        this.app = app;
        this.svg = null;
        this.viewport = null;
        this.backgroundLayer = null;
        this.cardLayer = null;
        this.streamLayer = null;
        this.themeLayer = null;
        this.teamLayer = null;
        this.logoLayer = null;
        this.zoom = null;
        this.width = 1200;
        this.height = 800;
        this.fitMinScale = 0.1;
        this.lastFitTransform = d3.zoomIdentity;
        this.snapToFitInProgress = false;
        this._memberCount = 0;
        this._lastLoggedK = null;
    }

    reset() {
        const { app } = this;
        const svgEl = document.getElementById('canvas');
        if (!svgEl) {
            console.error('canvas not found.');
            return;
        }

        this.width = svgEl.clientWidth || +svgEl.getAttribute('width') || 1200;
        this.height = svgEl.clientHeight || +svgEl.getAttribute('height') || 800;

        d3.select(svgEl).selectAll('*').remove();

        this.svg = d3.select(svgEl)
            .attr('width', this.width)
            .attr('height', this.height)
            .attr('cursor', 'grab');

        this.svgDefs = this.svg.append('defs');
        this.viewport = this.svg.append('g').attr('id', 'viewport');
        this.streamLayer = this.viewport.append('g').attr('id', 'streamLayer');
        this.themeLayer = this.viewport.append('g').attr('id', 'themeLayer');
        this.teamLayer = this.viewport.append('g').attr('id', 'teamLayer');
        this.cardLayer = this.viewport.append('g').attr('id', 'cardLayer');
        this.logoLayer = this.viewport.append('g').attr('id', 'logoLayer');

        this.zoom = d3.zoom()
            .filter((event) => {
                if (event.type === 'wheel') return !event.ctrlKey;
                if (event.type === 'mousedown') {
                    if (event.button !== 0) return false;
                    if (event.composedPath?.()?.some(el => el.tagName?.toLowerCase?.() === 'foreignobject')) return false;
                    if (app.interaction.mode === 'free-pan') return true;
                    if (app.interaction.mode === 'select') return false;
                    return false;
                }
                if (event.type.startsWith('touch')) return true;
                return !event.ctrlKey;
            })
            .scaleExtent([this.fitMinScale, ZOOM_MAX_SCALE])
            .on('start', (event) => {
                this.svg.classed('is-zooming', true);
                this.svg.attr('cursor', 'grabbing');
                if (!app.interaction.isDraggable && event?.sourceEvent?.type === 'mousedown') {
                    app.interaction.panMoved = false;
                    app.interaction.panStartPos = { x: event.sourceEvent.clientX, y: event.sourceEvent.clientY };
                }
                // Gesture-LOD: controlled by FEATURE_LOD build flag (off by default).
                // Set env var FEATURE_LOD=true before building to enable.
                // Activates only when zoomed far out (k < LOD_K_MAX) with a large dataset.
                if (__FEATURE_LOD__ && isMobileDevice() && event.sourceEvent !== null &&
                        this._memberCount > 100 && (event.transform?.k ?? 1) < LOD_K_MAX) {
                    this.svg.classed('lod-gesture-active', true);
                }
            })
            .on('zoom', (event) => {
                this.viewport.attr('transform', event.transform);
                this.onZoom?.();
                if (app.isAdvanced) {
                    const k = +event.transform.k.toFixed(3);
                    if (k !== this._lastLoggedK) {
                        this._lastLoggedK = k;
                        const lodActive = __FEATURE_LOD__ && this._memberCount > 100 && k < LOD_K_MAX;
                        console.log(`[solitaire] zoom k=${k} | LOD_K_MAX=${LOD_K_MAX} | members=${this._memberCount} | LOD=${lodActive ? '✓ active' : '✗ off'}`);
                    }
                }
                if (!app.interaction.isDraggable && event?.sourceEvent) {
                    const t = event.sourceEvent.type;
                    if (t === 'touchmove') {
                        app.interaction.panMoved = true;
                    } else if (t === 'mousemove' && app.interaction.panStartPos && !app.interaction.panMoved) {
                        const dx = event.sourceEvent.clientX - app.interaction.panStartPos.x;
                        const dy = event.sourceEvent.clientY - app.interaction.panStartPos.y;
                        if (dx * dx + dy * dy > PAN_MIN_MOVE_PX * PAN_MIN_MOVE_PX) {
                            app.interaction.panMoved = true;
                        }
                    }
                }
            })
            .on('end', (event) => {
                this.svg.classed('is-zooming', false);
                this.svg.classed('lod-gesture-active', false);
                this.svg.attr('cursor', 'grab');
                if (!app.interaction.isDraggable && app.interaction.panMoved) {
                    app.interaction.suppressClicksUntil = Date.now() + 250;
                }
                app.interaction.panMoved = false;
                app.interaction.panStartPos = null;
                if (isMobileDevice()) {
                    this._cullCards();
                }
            });

        this.svg.call(this.zoom);
        // Sync viewport transform with D3's cached __zoom so fitElementToView
        // reads consistent screen coordinates after re-render.
        this.viewport.attr('transform', d3.zoomTransform(this.svg.node()));
        this.installTrackpadPinchZoom(this.svg, this.zoom);

        const svgNode = this.svg.node();

        if (!window.__dsmGlobalContextMenuAttached) {
            window.__dsmGlobalContextMenuAttached = true;
            document.addEventListener('contextmenu', (e) => {
                const svgEl = document.getElementById('canvas');
                if (!svgEl || !svgEl.contains(e.target)) return;
                svgEl.style.touchAction = 'none';
                e.preventDefault();
                e.stopPropagation();
                app.contextMenu.show(e.clientX, e.clientY);
            }, true);
        }

        if (svgNode && !window.__panClickBlockerAttached) {
            window.__panClickBlockerAttached = true;
            svgNode.addEventListener('click', (e) => {
                if (!app.interaction.isDraggable && Date.now() < app.interaction.suppressClicksUntil) {
                    if (e.composedPath?.().some(el => el.tagName?.toLowerCase?.() === 'foreignobject')) return;
                    e.preventDefault();
                    e.stopImmediatePropagation();
                }
            }, true);
        }

        app.interaction.initSVGInteraction(svgNode);
    }

    installTrackpadPinchZoom(svgSel, zoomBehavior) {
        const svgNode = svgSel?.node?.();
        if (!svgNode || !zoomBehavior) return;
        if (svgNode.__pinchToD3Installed) return;
        svgNode.__pinchToD3Installed = true;

        const onWheel = (e) => {
            if (!e.ctrlKey) return;
            if (!svgNode.contains(e.target)) return;
            e.preventDefault();
            e.stopPropagation();
            const STEP = 1.12;
            const k = Math.sign(e.deltaY) > 0 ? 1 / STEP : STEP;
            const [x, y] = d3.pointer(e, svgNode);
            svgSel.call(zoomBehavior.scaleBy, k, [x, y]);
        };

        window.addEventListener('wheel', onWheel, { passive: false, capture: true });
    }

    fitToContent(paddingRatio = 0.9) {
        if (!this.viewport || !this.svg || !this.zoom) return;

        const bbox = this.viewport.node().getBBox();
        if (!bbox || !isFinite(bbox.width) || !isFinite(bbox.height) || bbox.width === 0 || bbox.height === 0) {
            this.fitMinScale = 0.1;
            this.lastFitTransform = d3.zoomIdentity;
            this.zoom.scaleExtent([this.fitMinScale, ZOOM_MAX_SCALE]);
            this.svg.call(this.zoom.transform, d3.zoomIdentity);
            return;
        }

        let scale = Math.min(this.width / bbox.width, this.height / bbox.height) * paddingRatio;
        scale = Math.min(scale, ZOOM_MAX_SCALE);

        const x = this.width / 2 - (bbox.x + bbox.width / 2) * scale;
        const y = this.height / 2 - (bbox.y + bbox.height / 2) * scale;

        const t = d3.zoomIdentity.translate(x, y).scale(scale);
        this.fitMinScale = scale;
        this.lastFitTransform = t;

        this.zoom.scaleExtent([this.fitMinScale, ZOOM_MAX_SCALE]);
        this.zoom.extent([[0, 0], [this.width, this.height]]);

        if (!this.app.interaction.isDraggable) {
            this.zoom.translateExtent([[-1e6, -1e6], [1e6, 1e6]]);
        } else {
            this.zoom.translateExtent([[bbox.x, bbox.y], [bbox.x + bbox.width, bbox.y + bbox.height]]);
        }

        this.svg.call(this.zoom.transform, t);
    }

    zoomToElement(element, desiredScale = 1.5, duration = 500) {
        if (!element || !this.svg) return;

        // On mobile, the element may be inside a culled (display:none) card group.
        // getBoundingClientRect() returns all-zeros for hidden elements, so un-hide first.
        if (isMobileDevice()) {
            const cardGroup = element.closest?.('g[data-cx]');
            if (cardGroup?.style.display === 'none') cardGroup.style.display = '';
        }

        const svgNode = this.svg.node();
        const t = d3.zoomTransform(svgNode);

        const elRect = element.getBoundingClientRect();
        const svgRect = svgNode.getBoundingClientRect();
        const centerScreenX = elRect.left + elRect.width / 2 - svgRect.left;
        const centerScreenY = elRect.top + elRect.height / 2 - svgRect.top;

        const [cx, cy] = t.invert([centerScreenX, centerScreenY]);

        const k = desiredScale;
        const offsetY = 190;
        const tx = this.width / 2 - cx * k;
        const ty = this.height / 2 - cy * k - offsetY;

        const targetTransform = d3.zoomIdentity.translate(tx, ty).scale(k);
        this.svg.transition().duration(duration).call(this.zoom.transform, targetTransform);

        const group = element.closest('g');
        if (group) highlightGroupUtils(d3.select(group));
    }

    fitElementToView(element, duration = 500) {
        if (!element || !this.svg || !this.zoom) return;

        // On mobile, card groups inside the container may be culled (display:none), making
        // getBoundingClientRect() return only the team-header bounds. Un-hide them first so
        // the scale is computed for the full element. _cullCards() re-evaluates on zoom.end.
        if (isMobileDevice()) {
            element.querySelectorAll?.('g[data-cx]').forEach(el => {
                if (el.style.display === 'none') el.style.display = '';
            });
        }

        const svgNode = this.svg.node();
        const t = d3.zoomTransform(svgNode);
        const elRect = element.getBoundingClientRect();
        const svgRect = svgNode.getBoundingClientRect();

        const centerScreenX = elRect.left + elRect.width  / 2 - svgRect.left;
        const centerScreenY = elRect.top  + elRect.height / 2 - svgRect.top;
        const [cx, cy] = t.invert([centerScreenX, centerScreenY]);

        const elW = elRect.width  / t.k;
        const elH = elRect.height / t.k;
        const k = Math.min(
            Math.min(this.width / elW, this.height / elH) * 0.88,
            ZOOM_MAX_SCALE
        );

        const tx = this.width  / 2 - cx * k;
        const ty = this.height / 2 - cy * k - 40;

        const targetTransform = d3.zoomIdentity.translate(tx, ty).scale(k);
        this.svg.transition().duration(duration).call(this.zoom.transform, targetTransform);

        highlightGroupUtils(d3.select(element));
    }

    bringToCorrectLayer(g) {
        const key = g.getAttribute('data-key') || '';
        if (key.startsWith('card::')) {
            this.cardLayer.node().appendChild(g);
        } else if (key.startsWith('team::')) {
            this.teamLayer.node().appendChild(g);
        } else if (key.startsWith('theme::')) {
            this.themeLayer.node().appendChild(g);
        } else if (key.startsWith('stream::')) {
            this.streamLayer.node().appendChild(g);
        }
    }

    getContentBBox() {
        // streamLayer always has visible rects (collapsed or not); cardLayer bbox is
        // zero when all cards are display:none (collapsed streams), so filter it out.
        const stream = this.streamLayer?.node()?.getBBox?.();
        const cards  = this.cardLayer?.node()?.getBBox?.();
        const boxes  = [stream, cards].filter(b => b && isFinite(b.width) && b.width > 0);
        if (!boxes.length) return null;
        const x1 = Math.min(...boxes.map(b => b.x));
        const y1 = Math.min(...boxes.map(b => b.y));
        const x2 = Math.max(...boxes.map(b => b.x + b.width));
        const y2 = Math.max(...boxes.map(b => b.y + b.height));
        return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 };
    }

    placeBrandLogo(maxWidth = 240, textMargin = 40) {
        if (!this.viewport || !this.logoLayer) return;

        const bbox = this.getContentBBox();
        if (!bbox) {
            console.warn('Visual outcome not found');
            return;
        }

        this.logoLayer.selectAll('*').remove();

        const scale = maxWidth / BRAND.logo.svgInline.viewBoxW;
        const height = Math.round(BRAND.logo.svgInline.viewBoxH * scale);
        const x = bbox.x + (bbox.width - maxWidth) / 2;
        const y = bbox.y + bbox.height + 600;

        this.logoLayer.append('path')
            .attr('d', BRAND.logo.svgInline.d)
            .attr('fill-rule', 'evenodd')
            .attr('clip-rule', 'evenodd')
            .attr('transform', `translate(${x}, ${y}) scale(${scale})`)
            .style('pointer-events', 'none');

        this.logoLayer.append('text')
            .attr('x', x + maxWidth / 2).attr('y', y + height + textMargin)
            .attr('text-anchor', 'middle')
            .attr('font-size', '24px')
            .attr('font-family', 'Arial, sans-serif')
            .attr('class', 'logo-tagline')
            .text(BRAND.tagline);

        if (!getQueryParam('search')) {
            this.fitToContent(0.9);
        }
    }

    makeResizable(group, rect, opts = {}) {
        const { app } = this;
        const minW = Number(opts.minWidth) || 200;
        const minH = Number(opts.minHeight) || 150;

        const title = group.select('text');

        const savedSize = app.scenario.getSavedSize(group);
        let w = (savedSize?.w ?? Number(rect.attr('width'))) || minW;
        let h = (savedSize?.h ?? Number(rect.attr('height'))) || minH;

        const handleSize = 14;
        const hitPad = 10;

        if (isMobileDevice()) {
            rect.attr('width', w).attr('height', h);
            if (title.node()) {
                const anchor = title.attr('text-anchor');
                if (anchor === 'middle') title.attr('x', w / 2);
            }
            if (typeof opts.onResize === 'function') opts.onResize({ width: w, height: h });
            return;
        }

        const handles = group.append('g').attr('class', 'resize-handles');
        handles.raise();

        const handleE  = handles.append('rect').attr('class', 'resize-handle e');
        const handleS  = handles.append('rect').attr('class', 'resize-handle s');
        const handleSE = handles.append('rect').attr('class', 'resize-handle se');
        const handleN  = handles.append('rect').attr('class', 'resize-handle n');
        const handleW  = handles.append('rect').attr('class', 'resize-handle w');
        const handleNW = handles.append('rect').attr('class', 'resize-handle nw');
        const handleNE = handles.append('rect').attr('class', 'resize-handle ne');
        const handleSW = handles.append('rect').attr('class', 'resize-handle sw');

        const hitE  = handles.append('rect').attr('class', 'resize-hit e');
        const hitS  = handles.append('rect').attr('class', 'resize-hit s');
        const hitSE = handles.append('rect').attr('class', 'resize-hit se');
        const hitN  = handles.append('rect').attr('class', 'resize-hit n');
        const hitW  = handles.append('rect').attr('class', 'resize-hit w');
        const hitNW = handles.append('rect').attr('class', 'resize-hit nw');
        const hitNE = handles.append('rect').attr('class', 'resize-hit ne');
        const hitSW = handles.append('rect').attr('class', 'resize-hit sw');

        const positionHandles = () => {
            const hitSize = handleSize + 2 * hitPad;

            handleE.attr('x', w - handleSize / 2).attr('y', h / 2 - handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleS.attr('x', w / 2 - handleSize / 2).attr('y', h - handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleSE.attr('x', w - handleSize / 2).attr('y', h - handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleN.attr('x', w / 2 - handleSize / 2).attr('y', -handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleW.attr('x', -handleSize / 2).attr('y', h / 2 - handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleNW.attr('x', -handleSize / 2).attr('y', -handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleNE.attr('x', w - handleSize / 2).attr('y', -handleSize / 2).attr('width', handleSize).attr('height', handleSize);
            handleSW.attr('x', -handleSize / 2).attr('y', h - handleSize / 2).attr('width', handleSize).attr('height', handleSize);

            hitE.attr('x', w - (handleSize / 2 + hitPad)).attr('y', h / 2 - (handleSize / 2 + hitPad)).attr('width', handleSize + 2 * hitPad).attr('height', handleSize + 2 * hitPad);
            hitS.attr('x', w / 2 - (handleSize / 2 + hitPad)).attr('y', h - (handleSize / 2 + hitPad)).attr('width', handleSize + 2 * hitPad).attr('height', handleSize + 2 * hitPad);
            hitSE.attr('x', w - (handleSize / 2 + hitPad)).attr('y', h - (handleSize / 2 + hitPad)).attr('width', handleSize + 2 * hitPad).attr('height', handleSize + 2 * hitPad);
            hitN.attr('x', w / 2 - hitSize / 2).attr('y', -hitSize / 2).attr('width', hitSize).attr('height', hitSize);
            hitW.attr('x', -hitSize / 2).attr('y', h / 2 - hitSize / 2).attr('width', hitSize).attr('height', hitSize);
            hitNW.attr('x', -hitSize / 2).attr('y', -hitSize / 2).attr('width', hitSize).attr('height', hitSize);
            hitNE.attr('x', w - hitSize / 2).attr('y', -hitSize / 2).attr('width', hitSize).attr('height', hitSize);
            hitSW.attr('x', -hitSize / 2).attr('y', h - hitSize / 2).attr('width', hitSize).attr('height', hitSize);
        };

        const applySize = () => {
            rect.attr('width', w).attr('height', h);
            if (!title.empty()) {
                const anchor = title.attr('text-anchor');
                if (anchor === 'middle') title.attr('x', w / 2);
            }
            positionHandles();
            if (typeof opts.onResize === 'function') opts.onResize({ width: w, height: h });
        };

        const makeDeltaTracker = () => {
            let prev = null;
            const getSvgPoint = (event) => {
                const t = d3.zoomTransform(this.svg.node());
                const [px, py] = d3.pointer(event, this.svg.node());
                return t.invert([px, py]);
            };
            return {
                start(event) { prev = getSvgPoint(event); },
                drag(event) {
                    const curr = getSvgPoint(event);
                    if (!prev) prev = curr;
                    const dx = curr[0] - prev[0];
                    const dy = curr[1] - prev[1];
                    prev = curr;
                    return { dx, dy };
                }
            };
        };

        const trackerE  = makeDeltaTracker();
        const trackerS  = makeDeltaTracker();
        const trackerSE = makeDeltaTracker();

        const applyTranslate = (dx, dy) => {
            const t = group.attr('transform') || 'translate(0,0)';
            const m = t.match(/translate\(([^,]+),\s*([^)]+)\)/);
            const x = m ? (+m[1] || 0) : 0;
            const y = m ? (+m[2] || 0) : 0;
            group.attr('transform', `translate(${x + dx},${y + dy})`);
        };

        const dragE = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerE.start(e); })
            .on('drag', (e) => { const { dx } = trackerE.drag(e); w = Math.max(minW, w + dx); applySize(); });

        const dragS = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerS.start(e); })
            .on('drag', (e) => { const { dy } = trackerS.drag(e); h = Math.max(minH, h + dy); applySize(); });

        const dragSE = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerSE.start(e); })
            .on('drag', (e) => { const { dx, dy } = trackerSE.drag(e); w = Math.max(minW, w + dx); h = Math.max(minH, h + dy); applySize(); });

        const dragN = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerS.start(e); })
            .on('drag', (e) => { const { dy } = trackerS.drag(e); const delta = Math.min(dy, h - minH); h -= delta; applyTranslate(0, delta); applySize(); });

        const dragW = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerE.start(e); })
            .on('drag', (e) => { const { dx } = trackerE.drag(e); const delta = Math.min(dx, w - minW); w -= delta; applyTranslate(delta, 0); applySize(); });

        const dragNW = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerSE.start(e); })
            .on('drag', (e) => { const { dx, dy } = trackerSE.drag(e); const dxC = Math.min(dx, w - minW); const dyC = Math.min(dy, h - minH); w -= dxC; h -= dyC; applyTranslate(dxC, dyC); applySize(); });

        const dragNE = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerSE.start(e); })
            .on('drag', (e) => { const { dx, dy } = trackerSE.drag(e); const dxC = Math.max(-(w - minW), dx); const dyC = Math.min(dy, h - minH); w = Math.max(minW, w + dxC); h -= dyC; applyTranslate(0, dyC); applySize(); });

        const dragSW = d3.drag()
            .on('start', (e) => { e.sourceEvent?.stopPropagation(); trackerSE.start(e); })
            .on('drag', (e) => { const { dx, dy } = trackerSE.drag(e); const dxC = Math.min(dx, w - minW); const dyC = Math.max(-(h - minH), dy); w -= dxC; h = Math.max(minH, h + dyC); applyTranslate(dxC, 0); applySize(); });

        handleE.call(dragE);  hitE.call(dragE);
        handleS.call(dragS);  hitS.call(dragS);
        handleSE.call(dragSE); hitSE.call(dragSE);
        handleN.call(dragN);  hitN.call(dragN);
        handleW.call(dragW);  hitW.call(dragW);
        handleNW.call(dragNW); hitNW.call(dragNW);
        handleNE.call(dragNE); hitNE.call(dragNE);
        handleSW.call(dragSW); hitSW.call(dragSW);

        handles.selectAll('.resize-handle, .resize-hit')
            .on('pointerdown', (event) => event.stopPropagation());

        handles
            .style('display', app.interaction.isDraggable ? null : 'none')
            .style('pointer-events', app.interaction.isDraggable ? 'all' : 'none');

        applySize();
    }

    // ─── Layout helpers (used inside render) ─────────────────────────────────

    _computeInARowForTeam(memberCount) {
        const BASE = 6;
        const TIER_SIZE = 18;
        const n = Math.max(0, Number(memberCount) || 0);
        if (n <= TIER_SIZE) return BASE;
        const tier = Math.floor((n - 1) / TIER_SIZE);
        return BASE * Math.pow(2, tier);
    }

    _uniqueMemberCount(members) {
        const arr = Array.isArray(members) ? members : [];
        const set = new Set();
        for (const m of arr) {
            const name = normalizeWs(m?.Name ?? m?.User ?? '').trim();
            if (name) set.add(name.toLowerCase());
        }
        return set.size || arr.length;
    }

    _computeTeamBoxWidth(teamInARow, memberWidth) {
        const BASE = 6;
        const extraOffsetX = teamInARow > 6 ? 80 + (100 * Math.floor(teamInARow / 24)) : 0;
        return (Number(teamInARow) || BASE) * (Number(memberWidth) || 0) + 100 + extraOffsetX;
    }

    _computeStreamWidthFromRows(layoutRows, secondLevelBoxPadX, leftPad = 60, rightPad = 60) {
        const rows = Array.isArray(layoutRows) ? layoutRows : [];
        const maxRow = rows.reduce((maxW, r) => {
            const themes = r?.themes || [];
            const themesWidth = themes.reduce((acc, t) => acc + (Number(t?.themeWidth) || 0), 0);
            const pads = themes.length > 1 ? (themes.length - 1) * (Number(secondLevelBoxPadX) || 0) : 0;
            return Math.max(maxW, leftPad + themesWidth + pads + rightPad);
        }, 0);
        return Math.max(600, maxRow);
    }

    // ─── Main render loop ─────────────────────────────────────────────────────

    render({ organizationWithManagers, filteredStreams, visibleStreamNames, headers, organization, streamBoosts = {}, themeBoosts = {}, teamBoosts = {} }) {
        const { app } = this;
        this._memberCount = 0;

        const dateValues = ['In team since'];
        const fieldsToShow = ['Role', BUSINESS_FUNCTION_FIELD, 'Company', 'Location', 'Room', ...dateValues];

        const nFields = fieldsToShow.length + 0.5;
        const rowHeight = 10;
        const memberWidth = 160, cardPad = 10, cardBaseHeight = nFields * 4 * rowHeight;
        const thirdLevelBoxPadX = 24;
        const secondLevelBoxPadX = 60;
        const firstLevelBoxPadY = 100;

        let streamY = 40;
        const streamX = 40;
        const collapsedKeys = this._getCollapsedKeys();

        const orderedStreams = Object.entries(organizationWithManagers)
            .sort(byBoost(streamBoosts));

        orderedStreams.forEach(([firstLevel, secondLevelItems]) => {
            if (firstLevel.includes(firstLevelNA)) return;

            if (filteredStreams) {
                const firstLevelNormalized = normalizeKey(firstLevel);
                if (!filteredStreams.has(firstLevel) && !filteredStreams.has(firstLevelNormalized)) return;
            }

            const firstLevelMembers =
                Object.values(organization[firstLevel] || {})
                    .flatMap(themeObj => Object.values(themeObj))
                    .flat();
            const firstLevelDescription =
                app.db.aggregateInfoByHeader(firstLevelMembers, headers, 'Team Stream Description', false, splitNarrativeValues)
                    ?.items?.join('\n\n') ?? '';

            let firstLevelBoxWidth = 600;

            const firstLevelGroup = this.streamLayer.append('g')
                .attr('class', 'draggable')
                .attr('transform', `translate(${streamX},${streamY})`)
                .attr('data-key', `stream::${normalizeKey(firstLevel)}`)
                .attr('data-description', firstLevelDescription || '');
            app.scenario.restoreGroupPosition(firstLevelGroup);

            const layoutRows = [];
            let currentRow = { themes: [], used: 0 };

            for (const [secondLevel, thirdLevelItems] of Object.entries(secondLevelItems).sort(byBoost(themeBoosts, `${firstLevel}::`))) {
                if (secondLevel.includes(secondLevelNA)) continue;

                const teamsMeta = Object.entries(thirdLevelItems || {}).sort(byBoost(teamBoosts, `${firstLevel}::${secondLevel}::`)).map(([thirdLevel, members]) => {
                    const memberCount = this._uniqueMemberCount(members);
                    const teamInARow = this._computeInARowForTeam(memberCount);
                    const teamRows = Math.max(1, Math.ceil(memberCount / teamInARow));
                    const teamBoxWidth = this._computeTeamBoxWidth(teamInARow, memberWidth);
                    return { thirdLevel, members, memberCount, teamInARow, teamRows, teamBoxWidth };
                });

                const nTeams = teamsMeta.length;

                if (currentRow.used > 0 && (currentRow.used + nTeams) > MAX_TEAMS_PER_ROW) {
                    layoutRows.push(currentRow);
                    currentRow = { themes: [], used: 0 };
                }

                const themeMaxRows = Math.max(1, ...teamsMeta.map(t => t.teamRows));
                const themeTeamsWidth = teamsMeta.reduce((acc, t) => acc + t.teamBoxWidth, 0);
                const themeInnerGaps = Math.max(0, nTeams - 1) * thirdLevelBoxPadX;
                const themeWidth = themeTeamsWidth + themeInnerGaps + SECOND_LEVEL_LABEL_EXTRA;

                currentRow.themes.push({ secondLevel, thirdLevelItems, teamsMeta, nTeams, themeMaxRows, themeWidth });
                currentRow.used += nTeams;
            }
            if (currentRow.themes.length) layoutRows.push(currentRow);

            firstLevelBoxWidth = this._computeStreamWidthFromRows(layoutRows, secondLevelBoxPadX);

            layoutRows.forEach(r => {
                r.rowMaxMemberRows = Math.max(1, ...r.themes.map(t => t.themeMaxRows));
                const cardsTopInTeam = 70 + 45;
                const TEAM_BOTTOM_PADDING = r.rowMaxMemberRows > 2 ? 60 : 40;
                r.teamBoxHeight =
                    cardsTopInTeam +
                    (r.rowMaxMemberRows - 1) * (cardBaseHeight + CARD_GAP_Y) +
                    cardBaseHeight +
                    TEAM_BOTTOM_PADDING;

                const dynamicBottomPadding = r.rowMaxMemberRows > 2 ? 60 : 40;
                r.themeBoxHeight = THEME_TOP_PADDING + r.teamBoxHeight + dynamicBottomPadding;
            });

            const streamKey = `stream::${normalizeKey(firstLevel)}`;
            const isCollapsed = collapsedKeys.has(streamKey);
            const fullHeight =
                layoutRows.reduce((acc, r) => acc + r.themeBoxHeight, 0) +
                (layoutRows.length > 1 ? (layoutRows.length - 1) * secondLevelRowPadY : 0) +
                175;
            const firstLevelBoxHeight = isCollapsed ? STREAM_COLLAPSED_HEIGHT : fullHeight;

            const isStreamCommunity = Object.values(organization[firstLevel] || {})
                .every(themeMap =>
                    Object.values(themeMap).every(teamMembers =>
                        teamMembers.some(m => !m.guestRole && m['Team Community'] === 'true')
                    )
                );

            const streamRect = firstLevelGroup.append('rect')
                .attr('class', isStreamCommunity ? 'stream-box stream-box--community' : 'stream-box')
                .attr('data-community', isStreamCommunity ? 'true' : null)
                .attr('width', firstLevelBoxWidth)
                .attr('height', firstLevelBoxHeight)
                .attr('data-full-height', fullHeight)
                .attr('rx', 40).attr('ry', 40);

            this.makeResizable(firstLevelGroup, streamRect, { minWidth: 600, minHeight: 300 });

            const titleY = isCollapsed ? STREAM_COLLAPSED_TITLE_Y : 105;
            const iconY  = isCollapsed ? STREAM_COLLAPSED_ICON_Y  : STREAM_ICON_Y;

            const titleText = firstLevelGroup.append('text')
                .attr('x', isStreamCommunity ? 130 : 50).attr('y', titleY)
                .attr('text-anchor', 'start')
                .attr('data-full-name', firstLevel)
                .attr('class', 'stream-title');

            titleText.text(firstLevel);

            if (isStreamCommunity) {
                firstLevelGroup.append('text')
                    .attr('class', 'community-badge community-badge--stream')
                    .attr('x', 50).attr('y', titleY)
                    .attr('text-anchor', 'start')
                    .attr('data-tooltip', communityTooltip("stream"))
                    .attr('aria-label', 'Community')
                    .text('🤝');
            }

            // SVG icon buttons — right-aligned, visible on stream hover.
            // Collapse/expand chevron (always present).
            const collapseIconX = firstLevelBoxWidth - 70;
            const collapseG = _appendIconG(
                firstLevelGroup,
                'stream-icon stream-icon--collapse',
                isCollapsed ? 'Expand stream' : 'Collapse stream',
                collapseIconX, iconY,
                ICON_CHEVRON_DOWN
            );
            collapseG.attr('transform', isCollapsed
                ? `translate(${collapseIconX}, ${iconY}) rotate(-90)`
                : `translate(${collapseIconX}, ${iconY}) rotate(180)`);
            collapseG.on('click', (e) => {
                e.stopPropagation();
                this.toggleCollapseStream(streamKey);
            });

            if (visibleStreamNames.length > 1) {
                _appendIconG(
                    firstLevelGroup,
                    'stream-icon stream-icon--hide',
                    'Hide this stream (ESC to reset)',
                    firstLevelBoxWidth - 140, iconY,
                    ICON_EYE_SLASH
                ).on('click', (e) => {
                    e.stopPropagation();
                    const key = normalizeKey(firstLevel);
                    const current = getAllowedStreamsSet();
                    let next;
                    if (!current) {
                        next = new Set(visibleStreamNames.map(s => normalizeKey(s)));
                        next.delete(key);
                    } else {
                        next = new Set(current);
                        next.delete(key);
                    }
                    app.setStreamFilter(next.size > 0 ? next : null);
                });

                _appendIconG(
                    firstLevelGroup,
                    'stream-icon stream-icon--isolate',
                    'Show this stream only (ESC to reset)',
                    firstLevelBoxWidth - 210, iconY,
                    ICON_EYE
                ).on('click', (e) => {
                    e.stopPropagation();
                    app.setStreamFilter(new Set([normalizeKey(firstLevel)]));
                });
            }

            if (firstLevelDescription !== '') {
                const openStreamDrawer = () => app.drawer.open({ name: firstLevel, description: firstLevelDescription, _permalinkSearch: `stream:${firstLevel}`, _showDetails: true });
                _appendIconG(
                    firstLevelGroup,
                    'stream-icon stream-icon--desc',
                    'View stream details',
                    firstLevelBoxWidth - 280, iconY,
                    ICON_INFO, ICON_INFO_DOT_CY
                ).on('click', (e) => {
                    e.stopPropagation();
                    openStreamDrawer();
                });

                firstLevelGroup.select('rect.stream-box')
                    .style('cursor', 'pointer')
                    .on('click', openStreamDrawer);
                firstLevelGroup.select('text.stream-title')
                    .style('cursor', 'pointer')
                    .on('click', openStreamDrawer);
            }

            let secondLevelYBase = streamY + 135;

            layoutRows.forEach((r) => {
                let secondLevelX = 60;
                const themeBoxHeightRow = r.themeBoxHeight;
                const teamBoxHeightRow  = r.teamBoxHeight;

                r.themes.forEach(({ secondLevel, thirdLevelItems, teamsMeta, nTeams, themeWidth }) => {
                    const secondLevelY = secondLevelYBase;

                    const originalThemeMembers = Object.values(organization[firstLevel]?.[secondLevel] || {}).flat();
                    const secondLevelDescription =
                        app.db.aggregateInfoByHeader(originalThemeMembers, headers, 'Team Theme Description', false, splitNarrativeValues)
                            ?.items?.join('\n\n') ?? '';

                    const secondLevelGroup = this.themeLayer.append('g')
                        .attr('class', 'draggable')
                        .attr('transform', `translate(${streamX + secondLevelX},${secondLevelY})`)
                        .attr('data-key', `theme::${normalizeKey(firstLevel)}::${normalizeKey(secondLevel)}`)
                        .attr('data-description', secondLevelDescription || '');
                    app.scenario.restoreGroupPosition(secondLevelGroup);

                    const isThemeCommunity = Object.values(organization[firstLevel]?.[secondLevel] || {})
                        .every(teamMembers => teamMembers.some(m => !m.guestRole && m['Team Community'] === 'true'));

                    const secondLevelRect = secondLevelGroup.append('rect')
                        .attr('class', isThemeCommunity ? 'theme-box theme-box--community' : 'theme-box')
                        .attr('data-community', isThemeCommunity ? 'true' : null)
                        .attr('width', themeWidth).attr('height', themeBoxHeightRow)
                        .attr('rx', 30).attr('ry', 30);
                    this.makeResizable(secondLevelGroup, secondLevelRect, { minWidth: 400, minHeight: 250 });

                    secondLevelGroup.append('text')
                        .attr('x', themeWidth / 2).attr('y', 85)
                        .attr('text-anchor', 'middle')
                        .attr('data-full-name', secondLevel)
                        .attr('class', 'theme-title')
                        .text(app.db.truncate(secondLevel));

                    if (isThemeCommunity) {
                        secondLevelGroup.append('text')
                            .attr('class', 'community-badge community-badge--theme')
                            .attr('x', 40).attr('y', 85)
                            .attr('text-anchor', 'start')
                            .attr('data-tooltip', communityTooltip("theme"))
                            .text('🤝');
                    }

                    if (secondLevelDescription !== '') {
                        const openThemeDrawer = () => app.drawer.open({ name: secondLevel, description: secondLevelDescription, _permalinkSearch: `theme:${secondLevel}`, _showDetails: true });
                        _appendIconG(
                            secondLevelGroup,
                            'theme-icon theme-icon--desc',
                            'View theme details',
                            themeWidth - 50, THEME_ICON_Y,
                            ICON_INFO, ICON_INFO_DOT_CY
                        ).on('click', (e) => {
                            e.stopPropagation();
                            openThemeDrawer();
                        });

                        secondLevelGroup.select('rect.theme-box')
                            .style('cursor', 'pointer')
                            .on('click', openThemeDrawer);
                        secondLevelGroup.select('text.theme-title')
                            .style('cursor', 'pointer')
                            .on('click', openThemeDrawer);
                    }

                    const effectiveTeamsMeta = (Array.isArray(teamsMeta) && teamsMeta.length)
                        ? teamsMeta
                        : Object.entries(thirdLevelItems || {}).map(([thirdLevel, members]) => {
                            const memberCount = Array.isArray(members) ? members.length : 0;
                            const tier = Math.floor((Math.max(0, memberCount) - 1) / 36);
                            const teamInARow = 6 * Math.pow(2, tier);
                            const teamRows = Math.max(1, Math.ceil(memberCount / teamInARow));
                            const teamBoxWidth = teamInARow * memberWidth + 100;
                            return { thirdLevel, members, teamInARow, teamRows, teamBoxWidth };
                        });

                    let teamOffsetX = 50;

                    effectiveTeamsMeta.forEach((tm) => {
                        const { thirdLevel, members, teamInARow, teamRows, teamBoxWidth } = tm;

                        const originalMembers = (organization[firstLevel]?.[secondLevel]?.[thirdLevel]) || [];
                        const services     = app.db.aggregateInfoByHeader(originalMembers, headers, 'Team Managed Services', true);
                        const description  = app.db.aggregateInfoByHeader(originalMembers, headers, 'Team Description', false, splitNarrativeValues)?.items?.join('\n\n') ?? '';
                        const channels     = app.db.aggregateInfoByHeader(originalMembers, headers, 'Team Channels', true)?.items;
                        const email        = app.db.aggregateInfoByHeader(originalMembers, headers, 'Team Email')?.items?.join('') ?? '';

                        const hasInfo = hasTeamDrawerContent({ description, services, channels, email });

                        const teamLocalX = teamOffsetX;
                        const teamLocalY = 130;

                        const cardsTopInTeam = 70 + 45 + 130 - teamLocalY;
                        const lastCardBottomInTeam = cardsTopInTeam + (teamRows - 1) * (cardBaseHeight + CARD_GAP_Y) + cardBaseHeight;
                        const teamBoxHeight = Math.max(TEAM_MIN_HEIGHT, lastCardBottomInTeam + TEAM_BOTTOM_PADDING_BASE);

                        const isCommunity = originalMembers.some(m => !m.guestRole && m['Team Community'] === 'true');

                        const thirdLevelGroup = this.teamLayer.append('g')
                            .attr('class', 'draggable')
                            .attr('transform', `translate(${streamX + secondLevelX + teamLocalX},${secondLevelY + teamLocalY})`)
                            .attr('data-key', `team::${normalizeKey(firstLevel)}::${normalizeKey(secondLevel)}::${normalizeKey(thirdLevel)}`)
                            .attr('data-community', isCommunity ? 'true' : null);
                        app.scenario.restoreGroupPosition(thirdLevelGroup);

                        const thirdLevelRect = thirdLevelGroup.append('rect')
                            .attr('class', isCommunity ? 'team-box team-box--community' : 'team-box')
                            .attr('width', teamBoxWidth).attr('height', teamBoxHeight)
                            .attr('rx', 20).attr('ry', 20);
                        this.makeResizable(thirdLevelGroup, thirdLevelRect, { minWidth: 360, minHeight: 220 });

                        const serviceCount = services?.items?.length || 0;

                        const teamTitle = thirdLevelGroup.append('text')
                            .attr('x', teamBoxWidth / 2).attr('y', 70)
                            .attr('text-anchor', 'middle')
                            .attr('data-full-name', thirdLevel)
                            .attr('data-services', services?.items?.filter(Boolean).join(', ') || '')
                            .attr('data-team-description', description || '')
                            .attr('data-team-email', email || '')
                            .attr('data-team-channels', JSON.stringify(channels || []))
                            .attr('class', 'team-title')
                            .text(app.db.truncate(thirdLevel));

                        if (serviceCount > 0) {
                            const bbox = teamTitle.node().getBBox();
                            const CHIP_H = 36, CHIP_ICON = 22, CHIP_PAD_L = 6, CHIP_GAP = 5, CHIP_PAD_R = 12;
                            const chipWidth = CHIP_PAD_L + CHIP_ICON + CHIP_GAP + String(serviceCount).length * 16 + CHIP_PAD_R;
                            const chipX = bbox.x + bbox.width + 10;
                            const chipY = 70 - CHIP_H + 5;

                            const chipG = thirdLevelGroup.append('g')
                                .attr('class', 'svc-chip')
                                .attr('cursor', 'pointer')
                                .attr('transform', `translate(${chipX},${chipY})`);

                            chipG.append('rect')
                                .attr('width', chipWidth).attr('height', CHIP_H)
                                .attr('rx', CHIP_H / 2);

                            chipG.append('image')
                                .attr('href', _dominoAccentIconSrc())
                                .attr('x', CHIP_PAD_L).attr('y', (CHIP_H - CHIP_ICON) / 2)
                                .attr('width', CHIP_ICON).attr('height', CHIP_ICON);

                            chipG.append('text')
                                .attr('x', CHIP_PAD_L + CHIP_ICON + CHIP_GAP)
                                .attr('y', CHIP_H * 0.72)
                                .attr('class', 'svc-chip__count')
                                .text(serviceCount);

                            chipG.on('click', (e) => {
                                e.stopPropagation();
                                app.drawer.open({ name: thirdLevel, description, elements: services, channels, email, elementsBaseUrl: (s) => `domino.html?search=id%3A"${encodeURIComponent(s)}"`, _permalinkSearch: `team:${thirdLevel}`, _showDetails: true, _openServices: true });
                            });
                        }

                        if (isCommunity) {
                            thirdLevelGroup.append('text')
                                .attr('class', 'community-badge community-badge--team')
                                .attr('x', 30).attr('y', 70)
                                .attr('text-anchor', 'start')
                                .attr('data-tooltip', communityTooltip("team"))
                                .text('🤝');
                        }

                        if (hasInfo) {
                            const openTeamDrawer = () => app.drawer.open({ name: thirdLevel, description, elements: services, channels, email, elementsBaseUrl: (s) => `domino.html?search=id%3A"${encodeURIComponent(s)}"`, _permalinkSearch: `team:${thirdLevel}`, _showDetails: true });

                            _appendIconG(
                                thirdLevelGroup,
                                'team-icon team-icon--desc',
                                'View team details',
                                teamBoxWidth - 50, TEAM_ICON_Y,
                                ICON_INFO, ICON_INFO_DOT_CY
                            ).on('click', (e) => {
                                e.stopPropagation();
                                openTeamDrawer();
                            });

                            thirdLevelGroup.select('rect.team-box').style('cursor', 'pointer').on('click', openTeamDrawer);
                            thirdLevelGroup.select('text.team-title').style('cursor', 'pointer').on('click', openTeamDrawer);
                        }

                        (members || []).forEach((member, mIdx) => {
                            this._memberCount++;
                            const col = mIdx % teamInARow;
                            const row = Math.floor(mIdx / teamInARow);

                            const cardX = 40 + secondLevelX + teamLocalX + 20 + col * (memberWidth + cardPad);
                            const cardY = secondLevelY + 70 + 45 + row * (cardBaseHeight + 10) + 130;

                            const group = this.cardLayer.append('g')
                                .attr('data-role', (member[ROLE_FIELD_WITH_MAPPING] || '').toString().trim())
                                .attr('data-company', (member[COMPANY_FIELD] || '').toString().trim())
                                .attr('data-location', (member[LOCATION_FIELD] || '').toString().trim())
                                .attr('data-function', (member[BUSINESS_FUNCTION_FIELD] || '').toString().trim())
                                .attr('data-room', (member['Room'] || '').toString().trim())
                                .attr('class', 'draggable')
                                .attr('transform', `translate(${cardX},${cardY})`)
                                .attr('data-cx', cardX)
                                .attr('data-cy', cardY)
                                .attr('data-key', `card::${normalizeKey(firstLevel)}::${normalizeKey(secondLevel)}::${normalizeKey(thirdLevel)}::${normalizeKey(member['Name'] || member['User'] || mIdx)}`);

                            const colorKey =
                                app.legend.colorBy === ROLE_FIELD_WITH_MAPPING ? group.attr('data-role') :
                                    app.legend.colorBy === COMPANY_FIELD ? group.attr('data-company') :
                                        app.legend.colorBy === BUSINESS_FUNCTION_FIELD ? group.attr('data-function') :
                                            group.attr('data-location');

                            app.legend.colorKeyMappings.set(
                                app.legend.colorBy,
                                (app.legend.colorKeyMappings.get(app.legend.colorBy) ?? new Set()).add(colorKey)
                            );

                            app.scenario.restoreGroupPosition(group);

                            const fill = app.legend.getCardFill(group) || NEUTRAL_COLOR;
                            const isNeutral = fill === NEUTRAL_COLOR;
                            const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
                            const effectiveFill = (isNeutral && isDark) ? '#1c1c1e' : fill;
                            const memberRect = group.append('rect')
                                .attr('class', 'profile-box')
                                .classed('profile-box--neutral', isNeutral)
                                .attr('width', memberWidth).attr('height', cardBaseHeight)
                                .attr('rx', 14).attr('ry', 14)
                                .attr('fill', effectiveFill);

                            if ((fill || '').toLowerCase() === '#ffffff' || fill === 'white') {
                                memberRect.attr('stroke', '#b8b8b8').attr('stroke-width', 1);
                            }

                            if (member.guestRole) {
                                memberRect.attr('stroke', '#333').attr('stroke-width', 1.5).attr('stroke-dasharray', '4 2');
                            }

                            // Capture card key synchronously so the async callback has a stable ID
                            const clipId = 'pc-' + String(group.attr('data-key') ?? mIdx)
                                .replace(/[^a-zA-Z0-9]/g, '-');

                            this._resolvePhoto(member[emailField]).then(photoPath => {
                                const photoSize = 60;
                                const photoX = (memberWidth - photoSize) / 2;
                                const photoY = 8;

                                const photoWrapper = group.append('g').attr('class', 'photo-wrapper');

                                if (photoPath === null) {
                                    const iconScale = photoSize / 64;
                                    const iconG = photoWrapper.append('g')
                                        .attr('class', 'avatar-placeholder')
                                        .attr('transform', `translate(${photoX},${photoY}) scale(${iconScale})`)
                                        .style('pointer-events', 'none');
                                    iconG.append('circle').attr('cx', 32).attr('cy', 20).attr('r', 10);
                                    iconG.append('path').attr('d', 'M16 52c0-8 8-14 16-14s16 6 16 14');
                                } else {
                                // Circular clip-path via objectBoundingBox so coordinates are
                                // independent of the card's position in the SVG coordinate space
                                this.svgDefs.append('clipPath')
                                    .attr('id', clipId)
                                    .attr('clipPathUnits', 'objectBoundingBox')
                                    .append('circle').attr('cx', 0.5).attr('cy', 0.5).attr('r', 0.5);

                                photoWrapper.append('image')
                                    .attr('href', photoPath)
                                    .attr('x', photoX).attr('y', photoY)
                                    .attr('width', photoSize).attr('height', photoSize)
                                    .attr('preserveAspectRatio', 'xMidYMid slice')
                                    .attr('clip-path', `url(#${clipId})`)
                                    .attr('aria-label', member.Name || 'profile photo')
                                    .style('pointer-events', 'none');
                                }

                                let nTeams = 0;
                                try { nTeams = countTeamsForMemberInOrg(member, app.visibleOrg) || 0; } catch {}

                                if (nTeams > 1) {
                                    const badgeR = 10;
                                    const bx = photoX + photoSize - badgeR - 1;
                                    const by = photoY + photoSize - badgeR - 1;
                                    const tooltipText = `Focus shared across ${nTeams} teams. Click to browse them`;

                                    const badgeG = photoWrapper.append('g')
                                        .attr('class', 'multi-team-badge')
                                        .attr('transform', `translate(${bx},${by})`)
                                        .style('cursor', 'pointer')
                                        .attr('role', 'button').attr('tabindex', 0)
                                        .attr('aria-label', tooltipText)
                                        .attr('data-tooltip', tooltipText);

                                    badgeG.append('circle').attr('r', badgeR).attr('fill', '#111').attr('stroke', '#fff').attr('stroke-width', 1.5);
                                    badgeG.append('text')
                                        .attr('text-anchor', 'middle').attr('dominant-baseline', 'central')
                                        .attr('fill', '#fff').style('font-weight', 600).style('font-size', `${badgeR + 2}px`)
                                        .text(nTeams);

                                    const triggerSearch = (e) => {
                                        e?.stopPropagation?.();
                                        const q = member.Name?.toLowerCase();
                                        if (q) app.search.search(q);
                                    };
                                    badgeG.on('click', triggerSearch);
                                    badgeG.on('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') triggerSearch(e); });
                                    badgeG.raise();
                                }
                            });

                            const nameY = 72;
                            const defaultNameBoxH = 24;
                            const nameFO = group.append('foreignObject')
                                .attr('x', 0).attr('y', nameY)
                                .attr('width', memberWidth).attr('height', defaultNameBoxH);
                            const nameDiv = nameFO.append('xhtml:div')
                                .attr('class', 'profile-name').html(member['Name']);

                            const adjustNameAndInfoHeights = () => {
                                const measured = nameDiv.node()?.scrollHeight || defaultNameBoxH;
                                const nameBoxH = Math.max(defaultNameBoxH, Math.ceil(measured) + 6);
                                nameFO.attr('height', nameBoxH);
                                const infoStartY = nameY + nameBoxH + 4;
                                const infoFOExisting = group.select('foreignObject .info').node()
                                    ? d3.select(group.select('foreignObject .info').node().closest('foreignObject'))
                                    : null;
                                if (infoFOExisting) infoFOExisting.attr('y', infoStartY);
                            };
                            requestAnimationFrame(() => requestAnimationFrame(adjustNameAndInfoHeights));

                            const infoDivFO_Y = nameY + defaultNameBoxH + 4;
                            const infoDiv = group.append('foreignObject')
                                .attr('x', 8).attr('y', infoDivFO_Y)
                                .attr('width', memberWidth - 16)
                                .attr('height', Math.max(0, cardBaseHeight - (infoDivFO_Y - 8)))
                                .append('xhtml:div').attr('class', 'info');

                            const memberEmail = member[emailField];
                            const isWebKit = /AppleWebKit/i.test(navigator.userAgent)
                                && /Safari/i.test(navigator.userAgent)
                                && !/(Chrome|Chromium|Edg)/i.test(navigator.userAgent);
                            const useSvgFabs = isWebKit
                                || /iPad|iPhone|iPod/i.test(navigator.userAgent)
                                || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

                            const photoSize = 60;
                            const photoX = (memberWidth - photoSize) / 2;
                            const photoY = 8;
                            const spacingX = 17;
                            const leftSpacingX = 1;
                            const fabSize = useSvgFabs ? 28 : 24;
                            const gap = useSvgFabs ? 3 : 8;
                            const fabsHeight = (fabSize * 2) + gap;
                            const rightX = Math.round(photoX + photoSize + spacingX);
                            const leftX = Math.round(photoX - spacingX - fabSize - leftSpacingX);
                            const fabsY = Math.round(photoY + Math.round((photoSize - fabsHeight) / 2) - 4);
                            const r = fabSize / 2;
                            const cx = Math.round(rightX + fabSize / 2);
                            const cy = Math.round(fabsY + fabSize / 2);
                            const dy = fabSize + gap;
                            const lc = { cx: Math.round(leftX + fabSize / 2), cy: Math.round(fabsY + fabSize / 2), r };

                            const reportClickHandler = (event) => {
                                event?.preventDefault?.();
                                event?.stopPropagation?.();
                                window.open(BRAND.urls.reportChange, '_blank', 'noopener');
                            };

                            const keringTalentUrl = getTalentKeringUrl(member);
                            const isInternal = app.db.isInternalCompany(member);

                            if (useSvgFabs) {
                                const reportG = group.append('g')
                                    .attr('class', 'contact-fabs-svg contact-fabs--left')
                                    .attr('transform', `translate(${lc.cx},${lc.cy})`);

                                const reportA = reportG.append('a')
                                    .attr('href', '#').attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                    .attr('class', 'contact-fab report')
                                    .attr('data-tooltip', 'Report change').attr('aria-label', 'Report change');

                                if (isInternal) {
                                    const talentA_left = reportG.append('a')
                                        .attr('href', keringTalentUrl).attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                        .attr('class', 'contact-fab kering-talent')
                                        .attr('data-tooltip', 'Kering Talent').attr('aria-label', 'Kering Talent');
                                    const talentG_left = talentA_left.append('g').attr('transform', `translate(0, ${dy})`);
                                    talentG_left.append('circle').attr('r', lc.r).attr('class', 'fab-circle');
                                    talentG_left.append('text').attr('class', 'fab-emoji').attr('text-anchor', 'middle').attr('dominant-baseline', 'central').text('👤');
                                    talentA_left.on('pointerdown', (e) => e.stopPropagation()).on('touchstart', (e) => e.stopPropagation());
                                }

                                const reportBtn = reportA.append('g').attr('transform', 'translate(0,0)');
                                reportBtn.append('circle').attr('r', lc.r).attr('class', 'fab-circle');
                                reportBtn.append('text').attr('class', 'fab-emoji').attr('text-anchor', 'middle').attr('dominant-baseline', 'central').text('📝');
                                reportA.on('pointerdown', (e) => e.stopPropagation()).on('touchstart', (e) => e.stopPropagation()).on('click', reportClickHandler);

                                if (memberEmail) {
                                    const fabsG = group.append('g')
                                        .attr('class', 'contact-fabs-svg contact-fabs--right')
                                        .attr('transform', `translate(${cx},${cy})`);

                                    const chatA = fabsG.append('a')
                                        .attr('href', `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(memberEmail)}`)
                                        .attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                        .attr('class', 'contact-fab chat').attr('data-tooltip', 'Chat').attr('aria-label', 'Chat');
                                    const chatG = chatA.append('g').attr('transform', 'translate(0,0)');
                                    chatG.append('circle').attr('r', r).attr('class', 'fab-circle');
                                    chatG.append('text').attr('class', 'fab-emoji').attr('text-anchor', 'middle').attr('dominant-baseline', 'central').text('💬');

                                    const mailA = fabsG.append('a')
                                        .attr('href', createOutlookUrl([memberEmail]))
                                        .attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                        .attr('class', 'contact-fab mail').attr('data-tooltip', 'Send email').attr('aria-label', 'Send email');
                                    const mailG = mailA.append('g').attr('transform', `translate(0, ${dy})`);
                                    mailG.append('circle').attr('r', r).attr('class', 'fab-circle');
                                    mailG.append('text').attr('class', 'fab-emoji').attr('text-anchor', 'middle').attr('dominant-baseline', 'central').text('✉️');

                                    fabsG.selectAll('a.contact-fab')
                                        .on('pointerdown', (e) => e.stopPropagation())
                                        .on('touchstart', (e) => e.stopPropagation());
                                }
                            } else {
                                const leftColumnCount = isInternal ? 2 : 1;
                                const leftFabsHeight = (fabSize * leftColumnCount) + (gap * (leftColumnCount - 1));
                                const fabsLeft = group.append('foreignObject')
                                    .attr('x', leftX).attr('y', fabsY).attr('width', fabSize).attr('height', leftFabsHeight)
                                    .attr('pointer-events', 'all').style('overflow', 'visible')
                                    .append('xhtml:div').attr('class', 'contact-fabs contact-fabs--left');

                                fabsLeft.append('a')
                                    .attr('class', 'contact-fab report').attr('href', '#')
                                    .attr('data-tooltip', 'Report change').attr('aria-label', 'Report change')
                                    .html('<span class="icon" aria-hidden="true">📝</span>')
                                    .on('click', reportClickHandler);

                                if (isInternal) {
                                    fabsLeft.append('a')
                                        .attr('class', 'contact-fab kering-talent').attr('href', keringTalentUrl)
                                        .attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                        .attr('data-tooltip', 'Kering Talent').attr('aria-label', 'Kering Talent')
                                        .html('<span class="icon" aria-hidden="true">👤</span>');
                                }

                                if (memberEmail) {
                                    const fabs = group.append('foreignObject')
                                        .attr('x', rightX).attr('y', fabsY).attr('width', fabSize).attr('height', fabsHeight)
                                        .attr('pointer-events', 'all').style('overflow', 'visible')
                                        .append('xhtml:div').attr('class', 'contact-fabs contact-fabs--right');

                                    fabs.append('a')
                                        .attr('class', 'contact-fab chat')
                                        .attr('href', `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(memberEmail)}`)
                                        .attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                        .attr('data-tooltip', 'Chat').attr('aria-label', 'Chat')
                                        .html('<span class="icon" aria-hidden="true">💬</span>');

                                    fabs.append('a')
                                        .attr('class', 'contact-fab mail').attr('href', createOutlookUrl([memberEmail]))
                                        .attr('target', '_blank').attr('rel', 'noopener noreferrer')
                                        .attr('data-tooltip', 'Send email').attr('aria-label', 'Send email')
                                        .html('<span class="icon" aria-hidden="true">✉️</span>');
                                }
                            }

                            group.classed('card', true);
                            group.selectAll('.contact-fabs-svg, .contact-fabs').each(function() {
                                this.parentNode.appendChild(this);
                            });
                            app.wireFabsInteractions(group);

                            Object.entries(member).forEach(([key, value]) => {
                                if (fieldsToShow.includes(key) && value) {
                                    let finalValue = value;
                                    if (dateValues.includes(key)) {
                                        const parsed = new Date(value);
                                        if (!isNaN(parsed)) finalValue = formatMonthYear(parsed);
                                    }
                                    const fieldKey = key.toLowerCase().replace(/\s+/g, '-');
                                    const isRole = key === 'Role';
                                    const isFunction = key === 'Function';
                                    const el = infoDiv.append('div')
                                        .attr('class', `${fieldKey}-field${(isRole || isFunction) ? ' field--clickable' : ''}`);
                                    if (isRole) {
                                        el.html(`<strong>${key}:</strong> <span class="field__link">${finalValue}</span>`);
                                        // Use native addEventListener — D3 .on() is unreliable inside foreignObject HTML
                                        el.node().addEventListener('click', (e) => {
                                            e.stopPropagation();
                                            const roleInfo = app.db.roleDetailsMapping.get(finalValue) || {};
                                            app.drawer.open({
                                                name: finalValue,
                                                description: roleInfo.description || '',
                                                _permalinkSearch: `role:${finalValue}`,
                                            });
                                        });
                                    } else if (isFunction) {
                                        el.html(`<strong>${key}:</strong> <span class="field__link">${finalValue}</span>`);
                                        el.node().addEventListener('click', (e) => {
                                            e.stopPropagation();
                                            const fnInfo = app.db.functionDetailsMapping?.get(finalValue) || {};
                                            if (!fnInfo.description) return;
                                            app.drawer.open({
                                                name: finalValue,
                                                description: fnInfo.description,
                                                _permalinkSearch: `function:${finalValue}`,
                                            });
                                        });
                                    } else {
                                        el.html(`<strong>${key}:</strong> ${finalValue}`);
                                    }
                                }
                            });
                        });

                        teamOffsetX += teamBoxWidth + thirdLevelBoxPadX;
                    });

                    secondLevelX += themeWidth + secondLevelBoxPadX;
                });

                secondLevelYBase += themeBoxHeightRow + secondLevelRowPadY;
            });

            streamY += firstLevelBoxHeight + firstLevelBoxPadY;
        });

        requestAnimationFrame(() => {
            this.placeBrandLogo(800, 50);
        });

        this._restoreCollapsedStreams();
        this.fitToContent(0.9);
        app.interaction.applyDraggableToggleState();
        requestAnimationFrame(() => {
            app.legend.setMode(app.legend.colorBy);
        });
    }

    // ─── Collapse / expand ────────────────────────────────────────────────────

    _getCollapsedKeys() {
        try { return new Set(JSON.parse(localStorage.getItem('dsm-collapsed-v1') || '[]')); }
        catch { return new Set(); }
    }

    _persistCollapsedKey(streamKey, collapsed) {
        const keys = this._getCollapsedKeys();
        collapsed ? keys.add(streamKey) : keys.delete(streamKey);
        localStorage.setItem('dsm-collapsed-v1', JSON.stringify([...keys]));
    }

    _hasCustomStreamLayout() {
        try {
            const saved = JSON.parse(localStorage.getItem('dsm-layout-v1:default') || '{}');
            return Object.keys(saved).some(k => k.startsWith('stream::'));
        } catch { return false; }
    }

    _restoreCollapsedStreams() {
        this._getCollapsedKeys().forEach(key => {
            const streamGroup = this.svg.select(`g[data-key="${key}"]`);
            if (streamGroup.empty()) return;
            d3.selectAll(_contentSelector(key)).style('display', 'none');
            this._setStreamHeaderY(streamGroup, STREAM_COLLAPSED_TITLE_Y, STREAM_COLLAPSED_ICON_Y, true);
            streamGroup.classed('stream-collapsed', true);
        });
    }

    toggleCollapseStream(streamKey) {
        const isCollapsed = this.svg.select(`g[data-key="${streamKey}"]`).classed('stream-collapsed');
        this._persistCollapsedKey(streamKey, !isCollapsed);
        if (this._hasCustomStreamLayout()) {
            isCollapsed ? this._expandInPlace(streamKey) : this._collapseInPlace(streamKey);
        } else {
            this.app.loadAndRender(this.app.db.cachedCsvText);
        }
        // Defer fitElementToView one RAF so it runs after placeBrandLogo's RAF
        // (which calls fitToContent and would otherwise cancel the zoom transition).
        requestAnimationFrame(() => {
            const streamGroup = this.svg.select(`g[data-key="${streamKey}"]`);
            if (!streamGroup.empty()) this.fitElementToView(streamGroup.node(), 0);
        });
    }

    _setStreamHeaderY(streamGroup, titleY, iconY, collapseRotated) {
        streamGroup.select('text.stream-title').attr('y', titleY);
        streamGroup.selectAll('.stream-icon').each(function() {
            const g = d3.select(this);
            const cx = g.attr('data-icon-x') || '0';
            const isCollapse = g.classed('stream-icon--collapse');
            if (isCollapse) {
                g.attr('transform', collapseRotated
                    ? `translate(${cx}, ${iconY}) rotate(-90)`
                    : `translate(${cx}, ${iconY}) rotate(180)`);
                g.attr('data-tooltip', collapseRotated ? 'Expand stream' : 'Collapse stream')
                 .attr('aria-label',   collapseRotated ? 'Expand stream' : 'Collapse stream');
            } else {
                g.attr('transform', `translate(${cx}, ${iconY})`);
            }
        });
    }

    _collapseInPlace(streamKey) {
        const streamGroup = this.svg.select(`g[data-key="${streamKey}"]`);
        if (streamGroup.empty()) return;
        streamGroup.select('rect.stream-box').attr('height', STREAM_COLLAPSED_HEIGHT);
        d3.selectAll(_contentSelector(streamKey)).style('display', 'none');
        this._setStreamHeaderY(streamGroup, STREAM_COLLAPSED_TITLE_Y, STREAM_COLLAPSED_ICON_Y, true);
        streamGroup.classed('stream-collapsed', true);
    }

    _expandInPlace(streamKey) {
        const streamGroup = this.svg.select(`g[data-key="${streamKey}"]`);
        if (streamGroup.empty()) return;
        const rect = streamGroup.select('rect.stream-box');
        rect.attr('height', rect.attr('data-full-height'));
        d3.selectAll(_contentSelector(streamKey)).style('display', '');
        this._setStreamHeaderY(streamGroup, 70, STREAM_ICON_Y, false);
        streamGroup.classed('stream-collapsed', false);
    }

    collapseAll() {
        this.svg.selectAll('g[data-key^="stream::"]').each((d, i, nodes) => {
            this._persistCollapsedKey(nodes[i].getAttribute('data-key'), true);
        });
        if (this._hasCustomStreamLayout()) {
            this.svg.selectAll('g[data-key^="stream::"]').each((d, i, nodes) => {
                this._collapseInPlace(nodes[i].getAttribute('data-key'));
            });
        } else {
            this.app.loadAndRender(this.app.db.cachedCsvText);
            this.app.renderer.fitToContent(0.9);
        }
    }

    expandAll() {
        localStorage.removeItem('dsm-collapsed-v1');
        this.app.loadAndRender(this.app.db.cachedCsvText);
        this.app.renderer.fitToContent(0.9);
    }

    _cullCards() {
        if (!isMobileDevice()) return;
        if (!this.svg || !this.cardLayer) return;

        const svgNode = this.svg.node();
        if (!svgNode) return;

        const t = d3.zoomTransform(svgNode);
        const w = svgNode.clientWidth || this.width;
        const h = svgNode.clientHeight || this.height;
        const margin = 300;

        const x0 = (-t.x - margin) / t.k;
        const y0 = (-t.y - margin) / t.k;
        const x1 = (w - t.x + margin) / t.k;
        const y1 = (h - t.y + margin) / t.k;

        this.cardLayer.selectAll('g[data-cx]').each(function () {
            const cx = +this.getAttribute('data-cx');
            const cy = +this.getAttribute('data-cy');
            this.style.display = (cx >= x0 && cx <= x1 && cy >= y0 && cy <= y1) ? '' : 'none';
        });

        // Re-hide cards that belong to collapsed streams (culling above may have re-shown them).
        this._getCollapsedKeys().forEach(key => {
            const base = key.replace(/^stream::/, '');
            this.cardLayer.selectAll(`g[data-key^="card::${base}::"]`).style('display', 'none');
        });
    }

    _resolvePhoto(email, timeoutMs = isMobileDevice() ? 1500 : 4000) {
        const baseName = (email?.split('@')[0] || '').replace('-ext', '').replace('.', '-');
        const fileName = `./assets/photos/${baseName}`;
        const candidates = [`${fileName}.webp`, `${fileName}.jpg`, `${fileName}.png`, `${fileName}.jpeg`];

        const tryWithTimeout = (url) => new Promise((resolve, reject) => {
            const img = new Image();
            const timer = setTimeout(() => { img.onload = img.onerror = null; reject(new Error('timeout')); }, timeoutMs);
            img.onload = () => { clearTimeout(timer); resolve(url); };
            img.onerror = () => { clearTimeout(timer); reject(new Error('error')); };
            img.src = url;
        });

        return candidates
            .reduce((chain, url) => chain.catch(() => tryWithTimeout(url)), Promise.reject())
            .catch(() => null);
    }
}
