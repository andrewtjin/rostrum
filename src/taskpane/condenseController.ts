// The Condense & Shrink controller — the range-scoped sibling of RostrumController.
//
// It is deliberately SEPARATE from RostrumController (not bolted on): its state (per-device condense
// settings, last-shrink size) is unrelated to invisibility's keepColors/armed/pureWholeBody, and it
// talks to the SECOND document seam (`RangeScopedPort`), not the whole-body `WordPort`. Every method
// returns the same `OpOutcome` shape RostrumController uses, so the ribbon's `toResult` + Track-Changes
// UI are reused verbatim. Each op runs under the shared Track-Changes gate: read the active range, run
// the matching PURE transform, write it back — one read + one write round-trip (tiny, so no
// cancellation is needed).

import { CondenseOptions, CondenseSettings, RangeRead, RangeScopedPort } from "../core/types";
import { TrackChangesActiveError, withTrackChangesGate } from "../core/guards";
import { createOfficeWordPort, WordRunner } from "../core/officeWordPort";
import { loadCondenseSettings, saveCondenseSettings, StorageLike } from "../core/settings";
import { resolveNormalHalfPts, shrinkFragment, unshrinkFragment } from "../core/shrink";
import { condenseFragment, uncondenseFragment } from "../core/condense";
import { Logger, logger as rootLogger } from "../core/debug";
import { OpOutcome } from "./controller";

/** What the Condense pane renders. */
export interface CondenseStatus {
  settings: CondenseSettings;
  /**
   * The half-point size the last Shrink/Unshrink applied (null = reset to Normal), or undefined when no
   * Shrink has run this session. The pane shows it as the current shrink size.
   */
  lastShrinkHalfPts?: number | null;
}

export interface CondenseControllerOptions {
  /** Injected in tests; defaults (in the adapter) to `Word.run`. */
  runner?: WordRunner;
  /** Inject a fake RangeScopedPort directly in tests (bypasses the Office adapter). */
  port?: RangeScopedPort;
  /** Per-device settings cache; defaults to window.localStorage when present. */
  storage?: StorageLike | null;
  logger?: Logger;
  /** Injected clock for deterministic timing in tests; defaults to Date.now. */
  now?: () => number;
}

/** The result of one pure transform, fed to the shared range-op runner. */
interface TransformOutcome {
  xml: string;
  changed: boolean;
  message: string;
}

export class CondenseController {
  private readonly port: RangeScopedPort;
  private readonly storage: StorageLike | null;
  private readonly log: Logger;
  private readonly now: () => number;

  private settings: CondenseSettings;
  private lastShrinkHalfPts: number | null | undefined = undefined;
  /** True while a range op runs — blocks overlapping ops (a settings change mid-op could mis-apply). */
  private inFlight = false;

  constructor(options: CondenseControllerOptions = {}) {
    this.storage = options.storage ?? (typeof localStorage !== "undefined" ? localStorage : null);
    this.log = options.logger ?? rootLogger("condense");
    this.now = options.now ?? ((): number => Date.now());
    this.settings = loadCondenseSettings(this.storage ?? memoryStorage());
    // The Office adapter implements RangeScopedPort; tests inject a fake port instead.
    this.port = options.port ?? createOfficeWordPort({ runner: options.runner, logger: rootLogger("adapter") });
  }

  // ----- Status / settings --------------------------------------------------

  status(): CondenseStatus {
    return { settings: this.getSettings(), lastShrinkHalfPts: this.lastShrinkHalfPts };
  }

  /** A defensive copy of the current settings (so callers can't mutate internal state). */
  getSettings(): CondenseSettings {
    return {
      usePilcrows: this.settings.usePilcrows,
      retainParagraphs: this.settings.retainParagraphs,
      reversal: this.settings.reversal,
      shrinkParagraphMarks: this.settings.shrinkParagraphMarks,
      omissionPatterns: this.settings.omissionPatterns.map((p) => ({ ...p }))
    };
  }

  /** Update settings (merge a partial patch) and persist to the device cache. No-op mid-operation. */
  setSettings(patch: Partial<CondenseSettings>): void {
    if (this.inFlight) {
      this.log.warn("condense settings change ignored — an operation is in progress");
      return;
    }
    this.settings = { ...this.settings, ...patch };
    try {
      saveCondenseSettings(this.storage ?? memoryStorage(), this.settings);
    } catch (e) {
      this.log.caught("persisting condense settings failed (kept in memory)", e);
    }
  }

  get busy(): boolean {
    return this.inFlight;
  }

  // ----- Shrink -------------------------------------------------------------

  /** Cycle non-kept card text down one font size (keeping underline/highlight/cite/headings full). */
  async shrink(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runRangeOp("shrink", autoToggleTrackChanges, (read) => {
      const out = shrinkFragment(read.ooxml, {
        normalHalfPts: resolveNormalHalfPts(read.ooxml),
        outlineLevels: read.outlineLevels,
        omissionPatterns: this.settings.omissionPatterns,
        shrinkParagraphMarks: this.settings.shrinkParagraphMarks
      });
      let message: string;
      if (out.refusedHeading) {
        message = "Can only shrink card text, not headings.";
      } else if (!out.changed) {
        message = "Nothing to shrink here (already smallest, or all kept).";
      } else {
        this.lastShrinkHalfPts = out.appliedSizeHalfPts ?? null;
        message = `Shrank to ${formatSize(out.appliedSizeHalfPts)}.`;
      }
      return { xml: out.xml, changed: out.changed, message };
    });
  }

  /** Reverse Shrink: clear the explicit sizes on non-kept runs (back to the inherited Normal size). */
  async unshrink(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runRangeOp("unshrink", autoToggleTrackChanges, (read) => {
      const out = unshrinkFragment(read.ooxml, read.outlineLevels);
      if (out.changed) this.lastShrinkHalfPts = null;
      return {
        xml: out.xml,
        changed: out.changed,
        message: out.changed ? "Reset to Normal size." : "Nothing to unshrink (already at Normal size)."
      };
    });
  }

  // ----- Condense -----------------------------------------------------------

  /** Condense using the pane's current settings (mode + reversal). */
  async condense(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runCondense("condense", autoToggleTrackChanges, {
      usePilcrows: this.settings.usePilcrows,
      retainParagraphs: this.settings.retainParagraphs,
      reversal: this.settings.reversal
    });
  }

  /** Merge paragraphs with a visible 6pt pilcrow (¶) at each former break (Verbatim parity). */
  async condenseWithPilcrows(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runCondense("condenseWithPilcrows", autoToggleTrackChanges, {
      usePilcrows: true,
      retainParagraphs: false,
      reversal: "marker"
    });
  }

  /** Full-merge everything into one block (invisible reversible markers; reversal from settings). */
  async fullCondense(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runCondense("fullCondense", autoToggleTrackChanges, {
      usePilcrows: false,
      retainParagraphs: false,
      reversal: this.settings.reversal
    });
  }

  /** Keep paragraph structure; collapse whitespace and drop blank lines (reversal from settings). */
  async retainParagraphsCondense(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runCondense("retainParagraphsCondense", autoToggleTrackChanges, {
      usePilcrows: false,
      retainParagraphs: true,
      reversal: this.settings.reversal
    });
  }

  /** Reverse Condense: every reversible break marker becomes a paragraph break again. */
  async uncondense(autoToggleTrackChanges = false): Promise<OpOutcome> {
    return this.runRangeOp("uncondense", autoToggleTrackChanges, (read) => {
      const out = uncondenseFragment(read.ooxml);
      return {
        xml: out.xml,
        changed: out.changed,
        message: out.changed
          ? `Restored ${out.breaksRestored} paragraph break(s).`
          : "Nothing to uncondense here (no Rostrum condense markers found)."
      };
    });
  }

  // ----- Internals ----------------------------------------------------------

  /** Shared condense dispatch: build the message from the engine's structured result. */
  private runCondense(label: string, autoToggleTrackChanges: boolean, opts: CondenseOptions): Promise<OpOutcome> {
    return this.runRangeOp(label, autoToggleTrackChanges, (read) => {
      const out = condenseFragment(read.ooxml, opts);
      let message: string;
      if (!out.changed) {
        message = "Nothing to condense here.";
      } else if (opts.retainParagraphs) {
        message = `Condensed ${out.paragraphsScanned} paragraph(s); dropped ${out.boundariesMarked} blank line(s).`;
      } else {
        message = `Condensed ${out.paragraphsScanned} paragraph(s) into one (${out.boundariesMarked} break(s) marked).`;
      }
      return { xml: out.xml, changed: out.changed, message };
    });
  }

  /**
   * The shared range-op runner: under the Track-Changes gate, read the active range, run the pure
   * `transform`, and write it back only when it changed. Translates every failure into the uniform
   * `OpOutcome` (the Track-Changes gate surfaces as a `trackChanges` outcome the pane can prompt on).
   */
  private async runRangeOp(
    label: string,
    autoToggleTrackChanges: boolean,
    transform: (read: RangeRead) => TransformOutcome
  ): Promise<OpOutcome> {
    if (this.inFlight) {
      return { status: "error", message: "Another Rostrum operation is still running — let it finish first." };
    }
    this.inFlight = true;
    const t0 = this.now();
    try {
      const { result, toggled } = await withTrackChangesGate(this.port, autoToggleTrackChanges, async () => {
        const read = await this.port.readActiveRangeOoxml();
        const out = transform(read);
        if (out.changed) await this.port.replaceActiveRangeOoxml(out.xml);
        return out;
      });
      const tookMs = this.now() - t0;
      const message = toggled ? `${result.message} (Track Changes toggled + restored)` : result.message;
      this.log.info(`${label} ok`, { tookMs, changed: result.changed });
      return { status: "ok", message, tookMs, detail: result };
    } catch (e) {
      if (e instanceof TrackChangesActiveError) {
        this.log.warn(`${label} blocked by Track Changes`, { mode: e.mode });
        return { status: "trackChanges", mode: e.mode };
      }
      this.log.caught(`${label} failed`, e);
      return { status: "error", message: errMessage(e) };
    } finally {
      this.inFlight = false;
    }
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Human size label: a number → "8pt", null/undefined → "Normal". */
export function formatSize(halfPts: number | null | undefined): string {
  return halfPts == null ? "Normal" : `${halfPts / 2}pt`;
}

function errMessage(e: unknown): string {
  if (e instanceof Error) return e.message;
  if (e && typeof e === "object" && "message" in e) return String((e as { message: unknown }).message);
  return String(e);
}

/** A throwaway in-memory storage when no real one exists (e.g. a privacy-locked host). */
function memoryStorage(): StorageLike {
  const m = new Map<string, string>();
  return {
    getItem: (k) => m.get(k) ?? null,
    setItem: (k, v) => void m.set(k, v)
  };
}
