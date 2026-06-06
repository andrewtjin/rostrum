// The feature registry — the single place the suite is assembled. Pure and host-free
// (no React, no Office.js) so it is unit-tested in Node and safe to import from the
// ribbon runtime. The shell asks it "what features/commands exist, and which are usable
// on THIS host?"; it never hard-codes a tool. Adding a tool = one `register()` call.
import { FeatureSupport } from "../core/types";
import { FeatureCommand, RostrumFeature } from "./types";

export class FeatureRegistry {
  // Insertion-ordered list (drives launcher order) + an id index for O(1) lookup.
  private readonly ordered: RostrumFeature[] = [];
  private readonly byId = new Map<string, RostrumFeature>();
  // Guards against two features claiming the same command id — which would make the
  // ribbon association ambiguous (the manifest FunctionName maps to exactly one handler).
  private readonly commandIds = new Set<string>();

  /**
   * Register a feature. Throws on a duplicate feature id or a command-id collision —
   * these are programmer errors that must fail loudly at startup, not silently shadow.
   * Returns `this` for fluent chaining in `index.ts`.
   */
  register(feature: RostrumFeature): this {
    if (this.byId.has(feature.id)) {
      throw new Error(`Rostrum: duplicate feature id "${feature.id}".`);
    }
    for (const command of feature.commands) {
      if (this.commandIds.has(command.id)) {
        throw new Error(
          `Rostrum: duplicate command id "${command.id}" (feature "${feature.id}").`
        );
      }
      this.commandIds.add(command.id);
    }
    this.byId.set(feature.id, feature);
    this.ordered.push(feature);
    return this;
  }

  /** All registered features, in registration order. */
  all(): readonly RostrumFeature[] {
    return this.ordered;
  }

  /** Look up one feature by id (nav routing, dialog routing). */
  get(id: string): RostrumFeature | undefined {
    return this.byId.get(id);
  }

  /** Features whose capability gate passes on this host (the launcher's "usable now" set). */
  available(features: FeatureSupport): RostrumFeature[] {
    return this.ordered.filter((feature) => feature.isAvailable(features));
  }

  /** Every command across all features (used by the ribbon to associate handlers). */
  commands(): FeatureCommand[] {
    return this.ordered.flatMap((feature) => feature.commands);
  }

  /** Find a single command by id (e.g. to invoke it outside the ribbon). */
  command(id: string): FeatureCommand | undefined {
    for (const feature of this.ordered) {
      const found = feature.commands.find((command) => command.id === id);
      if (found) return found;
    }
    return undefined;
  }
}
