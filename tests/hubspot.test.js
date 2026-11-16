"use strict";

// Test formatStage — it's an internal function, so we reconstruct it for testing

describe("hubspot.js formatStage (reconstructed)", () => {
  function formatStage(stage) {
    if (!stage) return "Unknown";
    const map = {
      appointmentscheduled:  "Appointment Scheduled",
      qualifiedtobuy:        "Qualified to Buy",
      presentationscheduled: "Presentation Scheduled",
      decisionmakerboughtin: "Decision Maker Bought In",
      contractsent:          "Contract Sent",
      closedwon:             "Closed Won",
      closedlost:            "Closed Lost",
    };
    const normalized = stage.toLowerCase().replace(/[\s_-]/g, "");
    return map[normalized] || stage.replace(/[_-]/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
  }

  it("returns 'Unknown' for empty/null/undefined", () => {
    expect(formatStage("")).toBe("Unknown");
    expect(formatStage(null)).toBe("Unknown");
    expect(formatStage(undefined)).toBe("Unknown");
  });

  it("maps 'appointmentscheduled'", () => {
    expect(formatStage("appointmentscheduled")).toBe("Appointment Scheduled");
  });

  it("maps 'qualifiedtobuy'", () => {
    expect(formatStage("qualifiedtobuy")).toBe("Qualified to Buy");
  });

  it("maps 'presentationscheduled'", () => {
    expect(formatStage("presentationscheduled")).toBe("Presentation Scheduled");
  });

  it("maps 'decisionmakerboughtin'", () => {
    expect(formatStage("decisionmakerboughtin")).toBe("Decision Maker Bought In");
  });

  it("maps 'contractsent'", () => {
    expect(formatStage("contractsent")).toBe("Contract Sent");
  });

  it("maps 'closedwon'", () => {
    expect(formatStage("closedwon")).toBe("Closed Won");
  });

  it("maps 'closedlost'", () => {
    expect(formatStage("closedlost")).toBe("Closed Lost");
  });

  it("handles mixed case input", () => {
    expect(formatStage("ClosedWon")).toBe("Closed Won");
    expect(formatStage("CONTRACTSENT")).toBe("Contract Sent");
  });

  it("handles input with spaces", () => {
    expect(formatStage("closed won")).toBe("Closed Won");
    expect(formatStage("contract sent")).toBe("Contract Sent");
  });

  it("handles input with underscores", () => {
    expect(formatStage("closed_won")).toBe("Closed Won");
    expect(formatStage("appointment_scheduled")).toBe("Appointment Scheduled");
  });

  it("handles input with hyphens", () => {
    expect(formatStage("closed-won")).toBe("Closed Won");
    expect(formatStage("decision-maker-bought-in")).toBe("Decision Maker Bought In");
  });

  it("title-cases unknown stages", () => {
    expect(formatStage("custom_stage")).toBe("Custom Stage");
    expect(formatStage("my-pipeline-step")).toBe("My Pipeline Step");
  });

  it("handles single-word unknown stage", () => {
    expect(formatStage("negotiation")).toBe("Negotiation");
  });

  it("handles already-formatted stage", () => {
    expect(formatStage("Closed Won")).toBe("Closed Won");
  });
});

// ── moveDealStage tests (mock node-fetch) ────────────────────
jest.mock("node-fetch", () => jest.fn());

process.env.HUBSPOT_TOKEN = "test-hubspot-token";

function mockResponse(body, { ok = true, status = 200 } = {}) {
  return {
    ok,
    status,
    json: async () => body,
    text: async () => (typeof body === "string" ? body : JSON.stringify(body)),
  };
}

describe("hubspot.moveDealStage", () => {
  let hubspot;
  let mockFetch;

  beforeEach(() => {
    jest.resetModules();
    mockFetch = require("node-fetch");
    mockFetch.mockReset();
    hubspot = require("../server/hubspot");
  });

  it("rejects non-numeric dealId", async () => {
    await expect(hubspot.moveDealStage("abc", "stage-1")).rejects.toThrow("Invalid dealId");
    await expect(hubspot.moveDealStage("", "stage-1")).rejects.toThrow("Invalid dealId");
    await expect(hubspot.moveDealStage(null, "stage-1")).rejects.toThrow("Invalid dealId");
    await expect(hubspot.moveDealStage("12a", "stage-1")).rejects.toThrow("Invalid dealId");
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it("rejects empty stageId", async () => {
    await expect(hubspot.moveDealStage("123", "")).rejects.toThrow("Invalid stageId");
    await expect(hubspot.moveDealStage("123", null)).rejects.toThrow("Invalid stageId");
    await expect(hubspot.moveDealStage("123", 42)).rejects.toThrow("Invalid stageId");
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it("throws when fetched deal has no pipeline", async () => {
    mockFetch.mockResolvedValueOnce(mockResponse({
      id: "123",
      properties: { dealname: "Test Deal", dealstage: "s1", pipeline: null },
    }));

    await expect(hubspot.moveDealStage("123", "s1")).rejects.toThrow("Deal has no pipeline set");
  });

  it("throws when stageId not in pipeline", async () => {
    mockFetch.mockResolvedValueOnce(mockResponse({
      id: "123",
      properties: { dealname: "Test Deal", dealstage: "s1", pipeline: "p1" },
    }));
    mockFetch.mockResolvedValueOnce(mockResponse({
      results: [
        {
          id: "p1",
          label: "Sales Pipeline",
          stages: [
            { id: "s1", label: "New",       displayOrder: 0 },
            { id: "s2", label: "Qualified", displayOrder: 1 },
          ],
        },
      ],
    }));

    await expect(hubspot.moveDealStage("123", "nonexistent-stage"))
      .rejects.toThrow(/does not belong to pipeline/);
  });

  it("throws when pipeline id is unknown", async () => {
    mockFetch.mockResolvedValueOnce(mockResponse({
      id: "123",
      properties: { dealname: "Test Deal", dealstage: "s1", pipeline: "p-unknown" },
    }));
    mockFetch.mockResolvedValueOnce(mockResponse({
      results: [
        { id: "p1", label: "Sales Pipeline", stages: [{ id: "s1", label: "New", displayOrder: 0 }] },
      ],
    }));

    await expect(hubspot.moveDealStage("123", "s1")).rejects.toThrow(/Unknown pipeline/);
  });

  it("happy path: PATCHes the deal with the new stage and returns raw pipeline stageLabel", async () => {
    // 1) fetch deal
    mockFetch.mockResolvedValueOnce(mockResponse({
      id: "123",
      properties: { dealname: "Test Deal", dealstage: "s1", pipeline: "p1" },
    }));
    // 2) getPipelines
    mockFetch.mockResolvedValueOnce(mockResponse({
      results: [
        {
          id: "p1",
          label: "Sales Pipeline",
          stages: [
            { id: "s1", label: "Appointment Scheduled", displayOrder: 0 },
            { id: "s2", label: "Closed Won",            displayOrder: 1 },
          ],
        },
      ],
    }));
    // 3) PATCH deal
    mockFetch.mockResolvedValueOnce(mockResponse({
      id: "123",
      properties: { dealstage: "s2" },
    }));

    const result = await hubspot.moveDealStage("123", "s2");

    // stageLabel is the raw pipeline label from HubSpot — consistent with
    // getDeals/searchDeals which also resolve labels via the pipelines cache.
    expect(result).toEqual({
      dealId:     "123",
      stageId:    "s2",
      stageLabel: "Closed Won",
      pipelineId: "p1",
    });

    const patchCall = mockFetch.mock.calls[2];
    expect(patchCall[0]).toContain("/crm/v3/objects/deals/123");
    expect(patchCall[1].method).toBe("PATCH");
    expect(JSON.parse(patchCall[1].body)).toEqual({
      properties: { dealstage: "s2" },
    });
  });
});

describe("hubspot.getPipelines memoization", () => {
  let hubspot;
  let mockFetch;

  beforeEach(() => {
    jest.resetModules();
    mockFetch = require("node-fetch");
    mockFetch.mockReset();
    hubspot = require("../server/hubspot");
  });

  it("caches pipelines across calls within TTL", async () => {
    mockFetch.mockResolvedValue(mockResponse({
      results: [
        { id: "p1", label: "Sales", stages: [{ id: "s1", label: "New", displayOrder: 0 }] },
      ],
    }));

    const a = await hubspot.getPipelines();
    const b = await hubspot.getPipelines();
    const c = await hubspot.getPipelines();

    expect(mockFetch).toHaveBeenCalledTimes(1);
    expect(a).toBe(b);
    expect(b).toBe(c);
  });
});

describe("PATCH /api/hubspot/deal/:id/stage validation logic", () => {
  // Mirrors validation in server/index.js route handler
  function validatePatchStage(id, body) {
    if (!id || !/^\d+$/.test(id))    return { status: 400, error: "Invalid deal id" };
    const { stageId } = body || {};
    if (!stageId || typeof stageId !== "string") return { status: 400, error: "stageId is required" };
    if (stageId.length > 64) return { status: 400, error: "stageId too long" };
    return { status: 200 };
  }

  it("400 on non-numeric id", () => {
    expect(validatePatchStage("abc", { stageId: "s1" })).toEqual({ status: 400, error: "Invalid deal id" });
    expect(validatePatchStage("", { stageId: "s1" })).toEqual({ status: 400, error: "Invalid deal id" });
    expect(validatePatchStage("12a", { stageId: "s1" })).toEqual({ status: 400, error: "Invalid deal id" });
  });

  it("400 on missing stageId", () => {
    expect(validatePatchStage("123", {})).toEqual({ status: 400, error: "stageId is required" });
    expect(validatePatchStage("123", { stageId: "" })).toEqual({ status: 400, error: "stageId is required" });
    expect(validatePatchStage("123", { stageId: 42 })).toEqual({ status: 400, error: "stageId is required" });
    expect(validatePatchStage("123", null)).toEqual({ status: 400, error: "stageId is required" });
  });

  it("400 on stageId too long (>64 chars)", () => {
    const longStage = "x".repeat(65);
    expect(validatePatchStage("123", { stageId: longStage })).toEqual({ status: 400, error: "stageId too long" });
  });

  it("passes validation with valid id and stageId", () => {
    expect(validatePatchStage("123", { stageId: "s1" })).toEqual({ status: 200 });
    expect(validatePatchStage("123", { stageId: "x".repeat(64) })).toEqual({ status: 200 });
  });
});

describe("PATCH /api/hubspot/deal/:id/stage targeted cache invalidation", () => {
  // Mirrors logic in server/index.js
  function invalidateForDeal(cache, id) {
    for (const [key, entry] of cache) {
      const deals = entry?.data?.deals;
      if (Array.isArray(deals) && deals.some((d) => d && d.id === id)) {
        cache.delete(key);
      }
    }
  }

  it("deletes only entries whose deals contain the moved deal id", () => {
    const cache = new Map();
    cache.set("k1", { data: { deals: [{ id: "111" }, { id: "222" }] } });
    cache.set("k2", { data: { deals: [{ id: "333" }] } });
    cache.set("k3", { data: { deals: [] } });
    cache.set("k4", { data: {} }); // missing deals
    cache.set("k5", { data: { deals: [{ id: "222" }] } });

    invalidateForDeal(cache, "222");

    expect(cache.has("k1")).toBe(false);
    expect(cache.has("k2")).toBe(true);
    expect(cache.has("k3")).toBe(true);
    expect(cache.has("k4")).toBe(true);
    expect(cache.has("k5")).toBe(false);
  });

  it("is defensive against missing or malformed entries", () => {
    const cache = new Map();
    cache.set("k1", null);
    cache.set("k2", {});
    cache.set("k3", { data: null });
    cache.set("k4", { data: { deals: null } });
    cache.set("k5", { data: { deals: [null, undefined, { id: "999" }] } });

    expect(() => invalidateForDeal(cache, "999")).not.toThrow();
    expect(cache.has("k5")).toBe(false);
    expect(cache.has("k1")).toBe(true);
    expect(cache.has("k2")).toBe(true);
    expect(cache.has("k3")).toBe(true);
    expect(cache.has("k4")).toBe(true);
  });
});
