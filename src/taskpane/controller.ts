// The task pane's brain, with NO React in it.
//
// All the orchestration the pane needs — build the adapter, load device defaults +
// the document manifest, resolve keep-colors, run hide/re-hide/show-all/apply-styles,
// translate every failure (especially the Track-Changes gate and cancellation) into a
// plain discriminated outcome — lives here so it is unit-tested with the same fake
// runner the adapter uses. `useRostrum` is a thin React shell over this controller.

import {
  DeviceDefaults,
  FeatureSupport,
  ResolvedSettings,
  TrackChangesMode
} from "../core/types";
import { hide as engineHide, reHide as engineReHide, showAll as engineShowAll } from "../core/invisibility";
import { loadManifest } from "../core/manifest";
import {
  DEFAULT_KEEP_COLORS,
  loadDeviceDefaults,
  loadPureWholeBody,
  resolveSettings,
  saveDeviceDefaults,
  savePureWholeBody,
  StorageLike
} from "../core/settings";
import { TrackChangesActiveError } from "../core/guards";
import {
  CancelledError,
  CiteRepairCapablePort,
  CiteRepairResult,
  CommitStrategy,
  createOfficeWordPort,
  ProgressInfo,
  WordRunner
} from "../core/officeWordPort";
import { ensureRostrumStyles, EnsureStylesResult } from "../core/officeStyles";
import { Logger, logger as rootLogger } from "../core/debug";

/** What the pane renders about the current document. */
export interface ControllerStatus {
  armed: boolean;
  keepColors: string[];
  /** Pure whole-body (avenue ⑦) — the default fast+lossless Hide path. False = compatibility opt-out. */
  pureWholeBody: boolean;
}

/** Uniform, explicit outcome for every user action (no throwing for control flow). */
export type OpOutcome =
  | { status: "ok"; message: string; tookMs: number; detail?: unknown }
  | { status: "trackChanges"; mode: TrackChangesMode }
  | { status: "cancelled" }
  | { status: "error"; message: string };

export interface ControllerOptions {
  features: FeatureSupport;
  /** Injected in tests; defaults (in the adapter) to `Word.run`. */
  runner?: WordRunner;
  /** Device-default keep-color cache; defaults to window.localStorage when present. */
  storage?: StorageLike | null;
  logger?: Logger;
  /** Progress sink for the pane's progress bar. */
  onProgress?: (info: ProgressInfo) => void;
  /** Override the commit strategy (defaults to the safe per-paragraph path). */
  commitStrategy?: CommitStrategy;
  /**
   * Force pure-whole-body (avenue ⑦) on/off. PRODUCTION leaves this unset → it reads the per-device
   * flag, defaulting ON (the wet-confirmed fast+lossless path). Tests pass `false` to exercise the
   * proxy/per-paragraph path against proxy-outline fixtures (which ⑦ ignores).
   */
  pureWholeBody?: boolean;
  /** Injected clock for deterministic timing in tests; defaults to Date.now. */
  now?: () => number;
}

export class RostrumController {
  private port: CiteRepairCapablePort;
  private readonly features: FeatureSupport;
  private readonly storage: StorageLike | null;
  private readonly runner?: WordRunner;
  private readonly onProgress?: (info: ProgressInfo) => void;
  /** Explicit commit-strategy override (tests). When set it wins over performance mode. */
  private readonly explicitStrategy?: CommitStrategy;
  private readonly log: Logger;
  private readonly now: () => number;

  private settings: ResolvedSettings = { keepColors: new Set(DEFAULT_KEEP_COLORS) };
  private deviceDefaults: DeviceDefaults | null = null;
  private armed = false;
  /** Pure whole-body (avenue ⑦) — the DEFAULT Hide path (one getOoxml + one insertOoxml, fast +
   *  losslessly reversible, wet-confirmed 2026-06-05). A user can opt OUT (→ per-paragraph). Set in
   *  the constructor from the per-device flag (default ON) or a test override. */
  private pureWholeBody = true;
  /** True while a Hide / Re-hide / Show All / Apply-Styles op runs. Blocks a mid-op pure-whole-body
   *  toggle (which rebuilds the adapter) so the invariant is enforced in the tested layer, not
   *  just the React `disabled` attribute. */
  private inFlight = false;
  /** Flipped by `cancel()`; the adapter polls it between read chunks. */
  private cancelFlag = false;

  constructor(options: ControllerOptions) {
    this.features = options.features;
    this.storage = options.storage ?? (typeof localStorage !== "undefined" ? localStorage : null);
    this.runner = options.runner;
    this.onProgress = options.onProgress;
    this.explicitStrategy = options.commitStrategy;
    this.log = options.logger ?? rootLogger("pane");
    this.now = options.now ?? ((): number => Date.now());
    // Pure whole-body (⑦) is the default; a per-device opt-OUT (or a test override) can turn it off.
    this.pureWholeBody = options.pureWholeBody ?? loadPureWholeBody(this.storage ?? memoryStorage(), true);
    this.port = this.createPort();
  }

  /** Build the adapter. Pure whole-body (⑦) is the default Hide path; the opt-out drops to the
   *  styles-faithful per-paragraph commit. The whole-body READ is used by both. The returned port
   *  also exposes `repairCites` (Apply Styles), hence the `CiteRepairCapablePort` type. */
  private createPort(): CiteRepairCapablePort {
    return createOfficeWordPort({
      runner: this.runner,
      commitStrategy: this.explicitStrategy ?? "whole-body",
      // Avenue ⑦ (default): when on, the adapter reads+commits the whole body with no proxies/
      // alignment and ignores commitStrategy (the per-paragraph path isn't taken). Opt-out → off.
      pureWholeBody: this.pureWholeBody,
      // Outline level is read from `Word.Paragraph.outlineLevel` (numeric, 1-based,
      // WordApi 1.1) — paragraphs expose no enum form, so use the default base.
      outline: { numberBase: "oneBased" },
      onProgress: this.onProgress,
      cancel: { isCancelled: () => this.cancelFlag },
      logger: rootLogger("adapter")
    });
  }

  /** Load device defaults + the document manifest and resolve the active keep-set. */
  async init(): Promise<ControllerStatus> {
    const span = this.log.span("controller.init");
    try {
      this.deviceDefaults = this.loadDefaultsSafely();
      const manifest = await loadManifest(this.port);
      this.settings = resolveSettings(manifest, this.deviceDefaults);
      this.armed = manifest?.active ?? false;
      span.end({ armed: this.armed, keepColors: this.keepColors });
      return this.status();
    } catch (e) {
      // Close the span as a failure so the bug-report timeline isn't left with a
      // dangling ▶ controller.init, then surface the original error to the pane.
      span.fail(e);
      throw e;
    }
  }

  /**
   * Re-read the document manifest and re-resolve the active keep-set, so a Hide / Re-hide / Show
   * All triggered OUTSIDE this pane is reflected here — most visibly the green "Invisibility ON"
   * indicator. The ribbon commands (commands.ts) use a SEPARATE `RostrumController` instance, so
   * their changes land in the document manifest (the shared source of truth) but never in this
   * controller's in-memory `armed` flag until we re-read. Lightweight: re-reads only the document
   * manifest (NOT device defaults). A no-op that returns the current status while an operation is in
   * flight — that op owns the truth, and a concurrent read could race its atomic commit. Returns the
   * (possibly refreshed) status so the pane can update its React state.
   */
  async refreshFromDocument(): Promise<ControllerStatus> {
    if (this.inFlight) return this.status();
    const manifest = await loadManifest(this.port);
    this.settings = resolveSettings(manifest, this.deviceDefaults);
    this.armed = manifest?.active ?? false;
    return this.status();
  }

  /** Read device defaults, tolerating a storage layer that throws (private mode, quota). */
  private loadDefaultsSafely(): DeviceDefaults | null {
    try {
      return loadDeviceDefaults(this.storage ?? memoryStorage());
    } catch (e) {
      this.log.caught("reading device defaults failed — using the built-in keep-set", e);
      return null;
    }
  }

  status(): ControllerStatus {
    return {
      armed: this.armed,
      keepColors: this.keepColors,
      pureWholeBody: this.pureWholeBody
    };
  }

  get keepColors(): string[] {
    return [...this.settings.keepColors];
  }

  /** Request cancellation of the in-flight read phase (safe — nothing written yet). */
  cancel(): void {
    this.cancelFlag = true;
    this.log.info("cancellation requested");
  }

  /**
   * Update the keep-colors and persist them as this device's default. They take
   * effect on the next Hide/Re-hide (which writes them into the document manifest).
   */
  setKeepColors(colors: string[]): void {
    this.settings = resolveSettings({ active: this.armed, keepColors: colors, schemaVersion: 1 }, null);
    const defaults: DeviceDefaults = { keepColors: this.keepColors };
    this.deviceDefaults = defaults;
    saveDeviceDefaults(this.storage ?? memoryStorage(), defaults);
    this.log.info("keep-colors updated", { keepColors: this.keepColors });
  }

  /** Whether the pure-whole-body (avenue ⑦) Hide path is currently on (the default). */
  get pureWholeBodyOn(): boolean {
    return this.pureWholeBody;
  }

  /**
   * Set the pure-whole-body (avenue ⑦) Hide path. It is the DEFAULT (on); this is the per-device
   * OPT-OUT — passing `false` rebuilds the adapter to use the slower, maximally-compatible
   * per-paragraph commit instead. Persists the preference and takes effect on the next Hide /
   * Re-hide. No-op mid-operation.
   */
  setPureWholeBody(on: boolean): void {
    if (on === this.pureWholeBody) return;
    if (this.inFlight) {
      this.log.warn("pure-whole-body toggle ignored — an operation is in progress");
      return;
    }
    this.pureWholeBody = on;
    try {
      savePureWholeBody(this.storage ?? memoryStorage(), on);
    } catch (e) {
      this.log.caught("persisting pure-whole-body preference failed (kept in memory)", e);
    }
    this.port = this.createPort();
    this.log.info(`pure-whole-body ${on ? "ON (default — one getOoxml + one insertOoxml)" : "OFF (compatibility per-paragraph path)"}`);
  }

  /** Hide non-keeper body text. `autoToggleTrackChanges` retries past the TC gate. */
  async hide(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runMutation(
      "hide",
      () => engineHide(this.port, this.settings, { autoToggleTrackChanges }),
      () => {
        this.armed = true;
      }
    );
  }

  /** Re-derive over the whole document (catches newly typed/pasted text). */
  async reHide(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runMutation(
      "reHide",
      () => engineReHide(this.port, this.settings, { autoToggleTrackChanges }),
      () => {
        this.armed = true;
      }
    );
  }

  /** Reveal everything Rostrum hid and disarm the document. */
  async showAll(): Promise<OpOutcome> {
    return this.runMutation(
      "showAll",
      () => engineShowAll(this.port),
      () => {
        this.armed = false;
      }
    );
  }

  /**
   * Apply the Rostrum heading/cite SIZES + pocket box (gated on WordApi 1.5; reflows the doc) AND
   * repair mis-styled cites (re-apply the cite character style so they're kept on Hide). The two
   * phases are INDEPENDENT and resilient:
   *
   *   * Cite-repair needs only the hard floor (getOoxml/insertOoxml), NOT getStyles/Style.font, so
   *     it runs even when the host can't apply sizes — making Apply Styles useful on more hosts.
   *   * A failure in one phase never aborts the other; both outcomes are surfaced. A repaired cite
   *     count is reported even when sizing is unsupported, and the op is only a hard error when
   *     BOTH sizing is unsupported AND no cites were repaired (nothing useful happened).
   */
  async applyStyles(): Promise<OpOutcome> {
    this.resetCancel();
    this.inFlight = true;
    const t0 = this.now();
    try {
      // Phase 1: sizes/box (gated). Capture failure as a value so it can't abort phase 2.
      let styleRes: EnsureStylesResult | null = null;
      let styleErr: unknown = null;
      try {
        styleRes = await ensureRostrumStyles({
          features: this.features,
          runner: this.runner,
          logger: rootLogger("styles")
        });
      } catch (e) {
        styleErr = e;
        this.log.caught("applyStyles: size/box phase failed (continuing to cite-repair)", e);
      }

      // Phase 2: cite-repair (whole-body OOXML — independent of the style APIs).
      let citeRes: CiteRepairResult | null = null;
      let citeErr: unknown = null;
      try {
        citeRes = await this.port.repairCites();
      } catch (e) {
        citeErr = e;
        this.log.caught("applyStyles: cite-repair phase failed (continuing)", e);
      }

      const tookMs = this.now() - t0;
      const detail = { styles: styleRes, styleError: messageOrNull(styleErr), cites: citeRes };

      // Both phases failed outright → a single combined error (nothing happened).
      if (styleErr && citeErr) {
        return { status: "error", message: `Apply Styles failed. ${errMessage(styleErr)}` };
      }

      // Compose the message from whatever each phase reports. The size phase is "unsupported"
      // (host lacks WordApi 1.5) vs "applied N" vs "errored"; the cite phase is "repaired M".
      const parts: string[] = [];
      const sizingUnsupported = !!styleRes?.unsupported;
      if (styleRes && !styleRes.unsupported) {
        parts.push(
          `Applied ${styleRes.applied.length} style(s)` +
            (styleRes.skipped.length ? `, skipped ${styleRes.skipped.length} (see diagnostics)` : "")
        );
      } else if (styleErr) {
        parts.push("style sizes failed (see diagnostics)");
      } else if (sizingUnsupported) {
        parts.push("style sizes unsupported on this host");
      }
      if (citeRes) {
        parts.push(`repaired ${citeRes.paragraphsRepaired} cite(s)`);
      } else if (citeErr) {
        parts.push("cite-repair failed (see diagnostics)");
      }

      // Hard error ONLY when sizing was unsupported AND cite-repair did nothing useful
      // (nothing actionable on this host) — otherwise report what DID happen.
      if (sizingUnsupported && (!citeRes || citeRes.paragraphsRepaired === 0)) {
        if (citeErr) return { status: "error", message: errMessage(citeErr) };
        return {
          status: "error",
          message: "This host can't apply Rostrum styles (needs desktop Word 1.5+)."
        };
      }

      return { status: "ok", message: `${parts.join("; ")}.`, tookMs, detail };
    } catch (e) {
      return { status: "error", message: errMessage(e) };
    } finally {
      this.inFlight = false;
    }
  }

  /** Shared mutation runner: timing, cancel reset, and outcome translation. */
  private async runMutation(
    label: string,
    op: () => Promise<unknown>,
    onSuccess: () => void
  ): Promise<OpOutcome> {
    this.resetCancel();
    this.inFlight = true;
    const t0 = this.now();
    try {
      const detail = await op();
      onSuccess();
      const tookMs = this.now() - t0;
      this.log.info(`${label} ok`, { tookMs, detail });
      return { status: "ok", message: summarizeResult(label, detail), tookMs, detail };
    } catch (e) {
      if (e instanceof TrackChangesActiveError) {
        this.log.warn(`${label} blocked by Track Changes`, { mode: e.mode });
        return { status: "trackChanges", mode: e.mode };
      }
      if (e instanceof CancelledError) {
        this.log.info(`${label} cancelled`);
        return { status: "cancelled" };
      }
      this.log.caught(`${label} failed`, e);
      return { status: "error", message: errMessage(e) };
    } finally {
      this.inFlight = false;
    }
  }

  private resetCancel(): void {
    this.cancelFlag = false;
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface HideLikeResult {
  paragraphsScanned: number;
  paragraphsChanged: number;
  paragraphsSkipped: number;
  trackChangesToggled?: boolean;
}

/** A friendly one-liner for the status bar from an engine result. */
function summarizeResult(label: string, detail: unknown): string {
  const r = detail as Partial<HideLikeResult>;
  if (r && typeof r.paragraphsChanged === "number") {
    const skipped = r.paragraphsSkipped ? `, ${r.paragraphsSkipped} skipped (unparseable)` : "";
    const tc = r.trackChangesToggled ? " (Track Changes toggled + restored)" : "";
    const verb = label === "showAll" ? "Revealed" : "Hid";
    return `${verb} ${r.paragraphsChanged} of ${r.paragraphsScanned} paragraphs${skipped}${tc}.`;
  }
  return `${label} complete.`;
}

function errMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  if (e && typeof e === "object" && "message" in e) return String((e as { message: unknown }).message);
  return String(e);
}

/** The error message of a thrown value, or null when there was no error (for structured detail). */
function messageOrNull(e: unknown): string | null {
  return e == null ? null : errMessage(e);
}

/** A throwaway in-memory storage when no real one exists (e.g. a privacy-locked host). */
function memoryStorage(): StorageLike {
  const m = new Map<string, string>();
  return {
    getItem: (k) => m.get(k) ?? null,
    setItem: (k, v) => void m.set(k, v)
  };
}
