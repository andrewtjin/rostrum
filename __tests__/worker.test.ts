// Unit tests for the download-counter Worker handler (worker/src/handler.js).
//
// The handler is plain CommonJS (outside tsconfig's include, so `tsc` ignores
// it) and is required here directly — ts-jest runs in CommonJS, so this needs no
// ESM gymnastics. We exercise it with a fake KV namespace and a stubbed global
// `fetch`, asserting both the happy path and every failure path that must NOT
// break a user's download.
//
// NOTE: worker/ is outside jest.config.js's coverage globs (src/core + gdocs/src/
// core only), so these tests are the SOLE correctness gate for the handler — a
// branch left untested ships silently green. Every route/path is enumerated.

// eslint-disable-next-line @typescript-eslint/no-var-requires
const { handleRequest, COUNTER_KEY, GDOCS_KEY } = require("../worker/src/handler.js");

/**
 * A Map-backed stand-in for a Cloudflare KV namespace (get/put only).
 * - failPut / failGet model a total KV outage (every op rejects).
 * - failGetKey models a SINGLE key's get rejecting while the other succeeds —
 *   needed to prove /count degrades each counter independently (never NaN).
 */
function makeEnv(opts: { failPut?: boolean; failGet?: boolean; failGetKey?: string } = {}) {
  const store = new Map<string, string>();
  return {
    MANIFEST_ORIGIN: "https://example.test/manifest.xml",
    CODE_GS_ORIGIN: "https://example.test/gdocs/Code.gs",
    COUNTER: {
      get: (k: string) =>
        opts.failGet || (opts.failGetKey && k === opts.failGetKey)
          ? Promise.reject(new Error("kv get down"))
          : Promise.resolve(store.get(k) ?? null),
      put: (k: string, v: string) => {
        if (opts.failPut) return Promise.reject(new Error("kv put down"));
        store.set(k, v);
        return Promise.resolve();
      }
    },
    _store: store
  };
}

const GET = (path: string) => new Request("https://w.test" + path);

afterEach(() => {
  // Restore any fetch stub between tests.
  // @ts-expect-error — test-only reassignment of the global.
  delete global.fetch;
});

describe("download-counter Worker", () => {
  // ---- Word surface: manifest.xml -----------------------------------------
  test("GET /manifest.xml counts then serves the manifest as an attachment", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("<manifest/>", { status: 200 }));
    const env = makeEnv();

    const res = await handleRequest(GET("/manifest.xml"), env);

    expect(res.status).toBe(200);
    expect(res.headers.get("content-disposition")).toContain('attachment; filename="manifest.xml"');
    expect(res.headers.get("content-type")).toContain("xml");
    expect(await res.text()).toBe("<manifest/>");
    expect(env._store.get(COUNTER_KEY)).toBe("1");
    // Proxied the configured origin, not the default.
    expect(global.fetch).toHaveBeenCalledWith("https://example.test/manifest.xml", expect.anything());
  });

  test("counts RAW downloads — every hit increments (no dedup)", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("<manifest/>", { status: 200 }));
    const env = makeEnv();

    await handleRequest(GET("/manifest.xml"), env);
    await handleRequest(GET("/manifest.xml"), env);
    await handleRequest(GET("/"), env); // bare-root alias also counts

    expect(env._store.get(COUNTER_KEY)).toBe("3");
  });

  test("origin unreachable → 302 to the canonical copy, but still counted", async () => {
    global.fetch = jest.fn().mockRejectedValue(new Error("network down"));
    const env = makeEnv();

    const res = await handleRequest(GET("/manifest.xml"), env);

    expect(res.status).toBe(302);
    expect(res.headers.get("location")).toBe("https://example.test/manifest.xml");
    expect(env._store.get(COUNTER_KEY)).toBe("1"); // intent-to-download still tallied
  });

  test("origin returns non-2xx → also 302 fallback", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("nope", { status: 503 }));
    const res = await handleRequest(GET("/manifest.xml"), makeEnv());
    expect(res.status).toBe(302);
  });

  test("KV put failure must NOT break the download", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("<manifest/>", { status: 200 }));
    const env = makeEnv({ failPut: true });

    const res = await handleRequest(GET("/manifest.xml"), env);

    expect(res.status).toBe(200);
    expect(await res.text()).toBe("<manifest/>");
  });

  test("HEAD /manifest.xml returns attachment headers but does NOT count or fetch the body", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("<manifest/>", { status: 200 }));
    const env = makeEnv();

    const res = await handleRequest(
      new Request("https://w.test/manifest.xml", { method: "HEAD" }),
      env
    );

    expect(res.status).toBe(200);
    expect(res.headers.get("content-disposition")).toContain("attachment");
    expect(env._store.get(COUNTER_KEY)).toBeUndefined(); // not counted
    expect(global.fetch).not.toHaveBeenCalled(); // no wasted upstream pull
  });

  // ---- Google Docs surface: code.gs ---------------------------------------
  test("GET /code.gs counts the gdocs surface then serves Code.gs as an attachment", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("/* Code.gs */", { status: 200 }));
    const env = makeEnv();

    const res = await handleRequest(GET("/code.gs"), env);

    expect(res.status).toBe(200);
    expect(res.headers.get("content-disposition")).toContain('attachment; filename="Code.gs"');
    expect(res.headers.get("content-type")).toContain("javascript");
    expect(await res.text()).toBe("/* Code.gs */");
    expect(env._store.get(GDOCS_KEY)).toBe("1");
    // Proxied the configured Code.gs origin, not the manifest one.
    expect(global.fetch).toHaveBeenCalledWith("https://example.test/gdocs/Code.gs", expect.anything());
  });

  test("/code.gs origin unreachable → 302 to the canonical Code.gs, still counted", async () => {
    global.fetch = jest.fn().mockRejectedValue(new Error("network down"));
    const env = makeEnv();

    const res = await handleRequest(GET("/code.gs"), env);

    expect(res.status).toBe(302);
    expect(res.headers.get("location")).toBe("https://example.test/gdocs/Code.gs");
    expect(env._store.get(GDOCS_KEY)).toBe("1");
  });

  test("/code.gs origin returns non-2xx → also 302 fallback", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("nope", { status: 404 }));
    const res = await handleRequest(GET("/code.gs"), makeEnv());
    expect(res.status).toBe(302);
  });

  test("HEAD /code.gs returns attachment headers but does NOT count or fetch the body", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("/* Code.gs */", { status: 200 }));
    const env = makeEnv();

    const res = await handleRequest(new Request("https://w.test/code.gs", { method: "HEAD" }), env);

    expect(res.status).toBe(200);
    expect(res.headers.get("content-disposition")).toContain('filename="Code.gs"');
    expect(env._store.get(GDOCS_KEY)).toBeUndefined();
    expect(global.fetch).not.toHaveBeenCalled();
  });

  test("/code.gs KV put failure must NOT break the download", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("/* Code.gs */", { status: 200 }));
    const env = makeEnv({ failPut: true });

    const res = await handleRequest(GET("/code.gs"), env);

    expect(res.status).toBe(200);
    expect(await res.text()).toBe("/* Code.gs */");
  });

  test("the two counters are INDEPENDENT — /code.gs never touches the Word tally and vice-versa", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("x", { status: 200 }));
    const env = makeEnv();

    await handleRequest(GET("/code.gs"), env);
    expect(env._store.get(GDOCS_KEY)).toBe("1");
    expect(env._store.get(COUNTER_KEY)).toBeUndefined(); // Word tally untouched

    await handleRequest(GET("/manifest.xml"), env);
    expect(env._store.get(COUNTER_KEY)).toBe("1");
    expect(env._store.get(GDOCS_KEY)).toBe("1"); // gdocs tally unchanged by a Word hit
  });

  // ---- /count: the unified cross-platform read ----------------------------
  test("GET /count returns both per-surface tallies plus the unified total (CORS-open)", async () => {
    const env = makeEnv();
    env._store.set(COUNTER_KEY, "42");
    env._store.set(GDOCS_KEY, "8");

    const res = await handleRequest(GET("/count"), env);

    expect(res.status).toBe(200);
    expect(res.headers.get("access-control-allow-origin")).toBe("*");
    expect(await res.json()).toEqual({ downloads: 42, gdocs: 8, total: 50 });
  });

  test("GET /count is all-zero when never set", async () => {
    const res = await handleRequest(GET("/count"), makeEnv());
    expect(await res.json()).toEqual({ downloads: 0, gdocs: 0, total: 0 });
  });

  test("total == downloads + gdocs for arbitrary values", async () => {
    const env = makeEnv();
    env._store.set(COUNTER_KEY, "1000");
    env._store.set(GDOCS_KEY, "337");
    const body = (await (await handleRequest(GET("/count"), env)).json()) as {
      downloads: number;
      gdocs: number;
      total: number;
    };
    expect(body.total).toBe(body.downloads + body.gdocs);
  });

  test("KV get failure on /count degrades BOTH counters to 0, never throws", async () => {
    const res = await handleRequest(GET("/count"), makeEnv({ failGet: true }));
    expect(await res.json()).toEqual({ downloads: 0, gdocs: 0, total: 0 });
  });

  test("/count degrades PER KEY — one counter readable, the other down → total never NaN", async () => {
    const env = makeEnv({ failGetKey: GDOCS_KEY }); // gdocs read throws, downloads read fine
    env._store.set(COUNTER_KEY, "5");

    const res = await handleRequest(GET("/count"), env);
    const body = (await res.json()) as { downloads: number; gdocs: number; total: number };

    expect(body.downloads).toBe(5);
    expect(body.gdocs).toBe(0); // degraded independently
    expect(body.total).toBe(5); // 5 + 0, deterministic — not NaN
    expect(Number.isNaN(body.total)).toBe(false);
  });

  test("a GARBLED stored value reads as 0 (the parseInt `|| 0` guard), and the next bump self-heals to 1", async () => {
    // A non-numeric KV value (corruption, a stray manual edit) must not poison the
    // tally with NaN — readCount's `|| 0` neutralizes it. Promised by the docstring;
    // proven here since worker/ is outside the coverage glob.
    const env = makeEnv();
    env._store.set(COUNTER_KEY, "abc");

    const body = (await (await handleRequest(GET("/count"), env)).json()) as {
      downloads: number;
      total: number;
    };
    expect(body.downloads).toBe(0);
    expect(Number.isNaN(body.total)).toBe(false);

    // A subsequent download recomputes from 0 → writes a clean "1", not "NaN"/"abc1".
    global.fetch = jest.fn().mockResolvedValue(new Response("<manifest/>", { status: 200 }));
    await handleRequest(GET("/manifest.xml"), env);
    expect(env._store.get(COUNTER_KEY)).toBe("1");
  });

  // ---- method / routing edges ---------------------------------------------
  test("OPTIONS /count answers the CORS preflight (204)", async () => {
    const res = await handleRequest(
      new Request("https://w.test/count", { method: "OPTIONS" }),
      makeEnv()
    );
    expect(res.status).toBe(204);
    expect(res.headers.get("access-control-allow-origin")).toBe("*");
    expect(res.headers.get("access-control-allow-methods")).toContain("GET");
  });

  test("KV get failure during a download is swallowed — file still served", async () => {
    global.fetch = jest.fn().mockResolvedValue(new Response("<manifest/>", { status: 200 }));
    const env = makeEnv({ failGet: true });

    const res = await handleRequest(GET("/manifest.xml"), env);

    expect(res.status).toBe(200);
    expect(await res.text()).toBe("<manifest/>");
  });

  test("non-GET method is rejected (405)", async () => {
    const res = await handleRequest(
      new Request("https://w.test/manifest.xml", { method: "POST" }),
      makeEnv()
    );
    expect(res.status).toBe(405);
  });

  test("unknown path is 404", async () => {
    const res = await handleRequest(GET("/nope"), makeEnv());
    expect(res.status).toBe(404);
  });
});
