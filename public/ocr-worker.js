/**
 * Web Worker that executes OpenAI API calls off the main thread.
 * Mobile browsers throttle/suspend the main thread when the tab is hidden,
 * but Workers are significantly less affected.
 */

async function callOpenAI(apiUrl, apiKey, model, systemPrompt, userContent) {
  const body = {
    model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent },
    ],
  };

  const res = await fetch(apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  });

  const data = await res.json().catch(() => ({}));
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

self.addEventListener("message", async (e) => {
  const { id, apiUrl, apiKey, model, systemPrompt, userContent } = e.data;
  try {
    const result = await callOpenAI(apiUrl, apiKey, model, systemPrompt, userContent);
    self.postMessage({ id, result });
  } catch (err) {
    self.postMessage({ id, error: err.message || String(err) });
  }
});
