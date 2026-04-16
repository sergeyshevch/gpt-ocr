/**
 * Web Worker that executes OpenAI API calls off the main thread.
 * Mobile browsers throttle/suspend the main thread when the tab is hidden,
 * but Workers are significantly less affected.
 */

const RATE_LIMIT_MAX_RETRIES = 8;

function parseRetryDelay(res, data) {
  const header = res.headers.get("retry-after");
  if (header) {
    const secs = parseFloat(header);
    if (Number.isFinite(secs) && secs > 0) return secs * 1000;
  }
  const msg = data?.error?.message || "";
  const match = msg.match(/try again in (\d+(?:\.\d+)?)\s*(ms|s|sec|seconds?)/i);
  if (match) {
    const val = parseFloat(match[1]);
    const unit = match[2].toLowerCase();
    if (Number.isFinite(val)) {
      return unit === "ms" ? val : val * 1000;
    }
  }
  return null;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function callOpenAI(apiUrl, apiKey, model, systemPrompt, userContent) {
  const body = {
    model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent },
    ],
  };

  const payload = JSON.stringify(body);

  for (let attempt = 0; ; attempt++) {
    const res = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: payload,
      signal: AbortSignal.timeout(120_000),
    });

    const data = await res.json().catch(() => ({}));

    if (res.status === 429 && attempt < RATE_LIMIT_MAX_RETRIES) {
      const suggested = parseRetryDelay(res, data);
      const backoff = Math.min(60_000, (2 ** attempt) * 1000);
      const delay = suggested ? Math.max(suggested + 200, 500) : backoff;
      await sleep(delay);
      continue;
    }

    if (!res.ok) {
      const msg = data.error?.message || data.message || `HTTP ${res.status}`;
      throw new Error(msg);
    }

    const text = data.choices?.[0]?.message?.content;
    if (typeof text !== "string") {
      throw new Error("Unexpected API response shape.");
    }
    return text.trim();
  }
}

self.addEventListener("message", async (e) => {
  const { id, apiUrl, apiKey, model, systemPrompt, userContent } = e.data;
  try {
    const result = await callOpenAI(apiUrl, apiKey, model, systemPrompt, userContent);
    self.postMessage({ id, result });
  } catch (err) {
    self.postMessage({ id, error: err.message || String(err) });
  }
});
