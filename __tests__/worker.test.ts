// Unit tests for the download-counter Worker handler (worker/src/handler.js).
//
// The handler is plain CommonJS (outside tsconfig's include, so `tsc` ignores
// it) and is required here directly — ts-jest runs in CommonJS, so this needs no
// ESM gymnastics. We exercise it with a fake KV namespace and a stubbed global
// `fetch`, asserting both the happy path and every failure path that must NOT
// break a user's download.

// eslint-disable-next-line @typescript-eslint/no-var-requires
const { handleRequest, COUNTER_KEY } = require("../worker/src/handler.js");

/** A Map-backed stand-in for a Cloudflare KV namespace (get/put only). */
function makeEnv(opts: { failPut?: boolean; failGet?: boolean } = {}) {
  const store = new Map<string, string>();
  return {
    MANIFEST_ORIGIN: "https://example.test/manifest.xml",
    COUNTER: {
      get: (k: string) =>
        opts.failGet ? Promise.reject(new Error("kv get down")) : Promise.resolve(store.get(k) ?? null),
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

  test("GET /count returns the tally as CORS-open JSON", async () => {
    const env = makeEnv();
    env._store.set(COUNTER_KEY, "42");

    const res = await handleRequest(GET("/count"), env);

    expect(res.status).toBe(200);
    expect(res.headers.get("access-control-allow-origin")).toBe("*");
    expect(await res.json()).toEqual({ downloads: 42 });
  });

  test("GET /count is 0 when never set", async () => {
    const res = await handleRequest(GET("/count"), makeEnv());
    expect(await res.json()).toEqual({ downloads: 0 });
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

  test("KV get failure on /count degrades to 0, never throws", async () => {
    const res = await handleRequest(GET("/count"), makeEnv({ failGet: true }));
    expect(await res.json()).toEqual({ downloads: 0 });
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
