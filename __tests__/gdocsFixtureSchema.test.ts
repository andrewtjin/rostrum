// Fixture-reality rail (plan A11.ii). Hand-built fixtures are only worth
// anything if they look like what Google actually sends, so every committed
// doc fixture is validated against the OFFICIAL Docs API discovery schemas
// (committed as fixtures/gdocs/discovery-schema.json — the Document-relevant
// $ref closure), and then linted for the wire conventions hand-builders get
// wrong: trailing newlines inside final runs, body starting at index 1, and
// the omitted-zero rgb rule (real payloads never write a 0.0 channel).
//
// The validator is deliberately tiny and dependency-free: a structural walk
// over the discovery JSON-schema dialect (type/properties/additionalProperties/
// items/enum/$ref). It is STRICT about unknown properties — a fixture key the
// schema does not declare is exactly the kind of invented reality this suite
// exists to catch.

import * as fs from "fs";
import * as path from "path";
import { DOC_FIELDS_MASK } from "../gdocs/src/core/parse";

const FIXTURE_DIR = path.join(__dirname, "fixtures", "gdocs");

const DOC_FIXTURE_NAMES = fs
  .readdirSync(FIXTURE_DIR)
  .filter((f) => f.endsWith(".json") && f !== "discovery-schema.json")
  .sort();

function loadJson(name: string): unknown {
  return JSON.parse(fs.readFileSync(path.join(FIXTURE_DIR, name), "utf8")) as unknown;
}

// ---------------------------------------------------------------------------
// Discovery-schema validator
// ---------------------------------------------------------------------------

/** The subset of the discovery schema dialect the Docs schemas actually use. */
interface DiscoverySchema {
  type?: string;
  $ref?: string;
  properties?: Record<string, DiscoverySchema>;
  additionalProperties?: DiscoverySchema;
  items?: DiscoverySchema;
  enum?: string[];
}

interface DiscoveryFile {
  discoveryRevision: string;
  schemas: Record<string, DiscoverySchema>;
}

const discovery = loadJson("discovery-schema.json") as DiscoveryFile;

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/**
 * Walk `instance` against `schema`, appending human-readable problems to
 * `errors` (never throwing — a validator that dies on the first problem hides
 * the rest). Path strings make a failing fixture immediately locatable.
 */
function validate(instance: unknown, schema: DiscoverySchema, where: string, errors: string[]): void {
  if (schema.$ref !== undefined) {
    const target = discovery.schemas[schema.$ref];
    if (target === undefined) {
      errors.push(`${where}: unresolved $ref ${schema.$ref}`);
      return;
    }
    validate(instance, target, where, errors);
    return;
  }
  switch (schema.type) {
    case "object": {
      if (!isPlainObject(instance)) {
        errors.push(`${where}: expected object, got ${typeof instance}`);
        return;
      }
      for (const [key, value] of Object.entries(instance)) {
        const prop = schema.properties !== undefined ? schema.properties[key] : undefined;
        if (prop !== undefined) {
          validate(value, prop, `${where}.${key}`, errors);
        } else if (schema.additionalProperties !== undefined) {
          // Map-typed fields (namedRanges et al.) declare their value schema
          // here; arbitrary keys are legitimate.
          validate(value, schema.additionalProperties, `${where}.${key}`, errors);
        } else {
          errors.push(`${where}.${key}: unknown property`);
        }
      }
      return;
    }
    case "array": {
      if (!Array.isArray(instance)) {
        errors.push(`${where}: expected array, got ${typeof instance}`);
        return;
      }
      const items = schema.items;
      if (items !== undefined) instance.forEach((v, i) => validate(v, items, `${where}[${i}]`, errors));
      return;
    }
    case "string": {
      if (typeof instance !== "string") {
        errors.push(`${where}: expected string, got ${typeof instance}`);
        return;
      }
      if (schema.enum !== undefined && !schema.enum.includes(instance)) {
        errors.push(`${where}: "${instance}" not in enum`);
      }
      return;
    }
    case "integer": {
      if (typeof instance !== "number" || !Number.isInteger(instance)) {
        errors.push(`${where}: expected integer`);
      }
      return;
    }
    case "number": {
      if (typeof instance !== "number" || !Number.isFinite(instance)) {
        errors.push(`${where}: expected number`);
      }
      return;
    }
    case "boolean": {
      if (typeof instance !== "boolean") errors.push(`${where}: expected boolean`);
      return;
    }
    default:
      // The Docs schemas only use the types above; anything else means the
      // committed schema (or this validator) drifted — surface it, don't skip.
      errors.push(`${where}: unsupported schema type ${String(schema.type)}`);
  }
}

function validateDocument(instance: unknown): string[] {
  const errors: string[] = [];
  validate(instance, { $ref: "Document" }, "$", errors);
  return errors;
}

/** Generic key/value walk over raw fixture JSON for the lints below. */
function walk(v: unknown, visit: (key: string, value: unknown, where: string) => void, where = "$"): void {
  if (Array.isArray(v)) {
    v.forEach((item, i) => walk(item, visit, `${where}[${i}]`));
    return;
  }
  if (!isPlainObject(v)) return;
  for (const [key, value] of Object.entries(v)) {
    visit(key, value, `${where}.${key}`);
    walk(value, visit, `${where}.${key}`);
  }
}

// ---------------------------------------------------------------------------
// The committed discovery schema itself
// ---------------------------------------------------------------------------

describe("committed discovery schema", () => {
  it("records its upstream revision (refresh provenance)", () => {
    expect(discovery.discoveryRevision).toMatch(/^\d{8}$/);
  });

  it("contains every schema the engine's read path stands on", () => {
    for (const id of [
      "Document",
      "Tab",
      "TabProperties",
      "DocumentTab",
      "Body",
      "StructuralElement",
      "Paragraph",
      "ParagraphElement",
      "ParagraphStyle",
      "TextRun",
      "TextStyle",
      "OptionalColor",
      "Color",
      "RgbColor",
      "Dimension",
      "Table",
      "TableRow",
      "TableCell",
      "NamedRanges",
      "NamedRange",
      "Range",
      "NamedStyles",
      "NamedStyle"
    ]) {
      expect(discovery.schemas[id]).toBeDefined();
    }
  });
});

// ---------------------------------------------------------------------------
// Every fixture conforms to the official Document schema
// ---------------------------------------------------------------------------

describe("fixtures conform to the discovery Document schema", () => {
  it.each(DOC_FIXTURE_NAMES)("%s", (name) => {
    expect(validateDocument(loadJson(name))).toEqual([]);
  });
});

describe("validator failure paths (it must actually be able to fail)", () => {
  it("rejects an unknown property", () => {
    expect(validateDocument({ bogusKey: 1 })).toEqual(["$.bogusKey: unknown property"]);
  });

  it("rejects an enum violation", () => {
    const errors = validateDocument({
      body: { content: [{ paragraph: { paragraphStyle: { namedStyleType: "HEADING_9" }, elements: [] } }] }
    });
    expect(errors).toEqual(['$.body.content[0].paragraph.paragraphStyle.namedStyleType: "HEADING_9" not in enum']);
  });

  it("rejects wrong primitive types, reporting every problem (not just the first)", () => {
    const errors = validateDocument({
      revisionId: 42,
      body: { content: [{ startIndex: 1.5, paragraph: { elements: [{ textRun: { content: 7 } }] } }] }
    });
    expect(errors).toContain("$.revisionId: expected string, got number");
    expect(errors).toContain("$.body.content[0].startIndex: expected integer");
    expect(errors).toContain("$.body.content[0].paragraph.elements[0].textRun.content: expected string, got number");
    expect(errors).toHaveLength(3);
  });

  it("rejects a junk array where the schema expects one", () => {
    expect(validateDocument({ tabs: "nope" })).toEqual(["$.tabs: expected array, got string"]);
  });
});

// ---------------------------------------------------------------------------
// Fixture lints — the wire conventions hand-builders get wrong (plan A11.ii)
// ---------------------------------------------------------------------------

describe("fixture lint", () => {
  /** All paragraph payloads found anywhere in a fixture (body, cells, tabs). */
  function collectParagraphs(fixture: unknown): { where: string; para: Record<string, unknown> }[] {
    const out: { where: string; para: Record<string, unknown> }[] = [];
    walk(fixture, (key, value, where) => {
      if (key === "paragraph" && isPlainObject(value)) out.push({ where, para: value });
    });
    return out;
  }

  /** All body payloads (legacy top-level and per-documentTab). */
  function collectBodies(fixture: unknown): { where: string; body: Record<string, unknown> }[] {
    const out: { where: string; body: Record<string, unknown> }[] = [];
    walk(fixture, (key, value, where) => {
      if (key === "body" && isPlainObject(value)) out.push({ where, body: value });
    });
    return out;
  }

  it.each(DOC_FIXTURE_NAMES)("%s: every paragraph's text ends with its newline inside the final textRun", (name) => {
    const paragraphs = collectParagraphs(loadJson(name));
    expect(paragraphs.length).toBeGreaterThan(0);
    for (const { where, para } of paragraphs) {
      const elements = para.elements;
      if (!Array.isArray(elements) || elements.length === 0) throw new Error(`${where}: paragraph without elements`);
      const last = elements[elements.length - 1] as Record<string, unknown>;
      const run = last.textRun as Record<string, unknown> | undefined;
      // The Docs API has no standalone paragraph-mark node: the newline lives
      // in the FINAL text run, never inside a chip/object element.
      if (run === undefined || typeof run.content !== "string" || !run.content.endsWith("\n")) {
        throw new Error(`${where}: final element must be a textRun ending with \\n`);
      }
    }
  });

  it.each(DOC_FIXTURE_NAMES)("%s: every body segment opens at index 1 and closes on a paragraph", (name) => {
    const bodies = collectBodies(loadJson(name));
    expect(bodies.length).toBeGreaterThan(0);
    for (const { where, body } of bodies) {
      const content = body.content;
      if (!Array.isArray(content) || content.length === 0) throw new Error(`${where}: body without content`);
      const firstPara = content.find((se) => isPlainObject(se) && isPlainObject(se.paragraph)) as
        | Record<string, unknown>
        | undefined;
      if (firstPara === undefined) throw new Error(`${where}: body has no paragraph`);
      // Index 0 belongs to the (masked-out) section break; text starts at 1.
      expect(firstPara.startIndex).toBe(1);
      // Real bodies always END with a paragraph — a trailing table is invalid.
      const last = content[content.length - 1] as Record<string, unknown>;
      expect(isPlainObject(last.paragraph)).toBe(true);
    }
  });

  it.each(DOC_FIXTURE_NAMES)("%s: no explicit 0.0 rgb channel (real payloads omit zero channels)", (name) => {
    const violations: string[] = [];
    walk(loadJson(name), (key, value, where) => {
      if ((key === "red" || key === "green" || key === "blue") && value === 0) violations.push(where);
    });
    expect(violations).toEqual([]);
  });

  it("includes the surrogate-pair fixture, and it really carries surrogates", () => {
    expect(DOC_FIXTURE_NAMES).toContain("surrogate.json");
    const rawText = fs.readFileSync(path.join(FIXTURE_DIR, "surrogate.json"), "utf8");
    expect(/[\uD800-\uDBFF]/.test(rawText)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// DOC_FIELDS_MASK reality check — every identifier in the mask must exist as
// a property name somewhere in the official Document schema closure, so a
// typo'd selector (silently dropping a field from every read) cannot survive.
// ---------------------------------------------------------------------------

describe("DOC_FIELDS_MASK against the discovery schema", () => {
  it("selects only real property names", () => {
    const known = new Set<string>();
    for (const schema of Object.values(discovery.schemas)) {
      for (const prop of Object.keys(schema.properties ?? {})) known.add(prop);
    }
    const tokens = DOC_FIELDS_MASK.split(/[(),]+/).filter((t) => t.length > 0);
    expect(tokens.length).toBeGreaterThan(0);
    const unknown = tokens.filter((t) => !known.has(t));
    expect(unknown).toEqual([]);
  });
});
