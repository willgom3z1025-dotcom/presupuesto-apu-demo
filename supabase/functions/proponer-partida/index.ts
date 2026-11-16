// Supabase Edge Function: proponer-partida
// Recibe {prompt} desde la aplicación (GitHub Pages) y lo envía a Gemini (API de Google, plan gratuito).
// La llave vive solo aquí, como secreto del proyecto (GEMINI_API_KEY); nunca en el HTML.
// Secretos: GEMINI_API_KEY (obligatorio), GEMINI_MODEL (opcional: modelos separados por coma, en orden de preferencia).
// Si un modelo está saturado o no existe, se prueba el siguiente.
// En la aplicación: Proyecto > Inteligencia artificial > URL = https://<ref>.supabase.co/functions/v1/proponer-partida

const ORIGENES = ["https://willgom3z1025-dotcom.github.io"];
const MODELOS = (Deno.env.get("GEMINI_MODEL") ?? "gemini-3.8-flash,gemini-3.6-flash,gemini-3.5-flash-lite,gemini-3.1-flash-lite").split(",").map((m) => m.trim()).filter(Boolean);
const MAX_PROMPT = 60000; // caracteres; el catálogo estándar ocupa unos 20,000

function cors(origin: string | null) {
  const ok = origin && ORIGENES.includes(origin) ? origin : ORIGENES[0];
  return {
    "Access-Control-Allow-Origin": ok,
    "Access-Control-Allow-Headers": "authorization, apikey, content-type, x-client-info",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  };
}
const json = (h: Record<string, string>, body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...h, "Content-Type": "application/json" } });

Deno.serve(async (req) => {
  const h = cors(req.headers.get("origin"));
  if (req.method === "OPTIONS") return new Response("ok", { headers: h });
  if (req.method !== "POST") return json(h, { error: "Método no permitido" }, 405);
  const origen = req.headers.get("origin");
  if (!origen || !ORIGENES.includes(origen)) return json(h, { error: "Origen no permitido" }, 403);
  try {
    const { prompt } = await req.json();
    if (typeof prompt !== "string" || !prompt || prompt.length > MAX_PROMPT) return json(h, { error: "prompt inválido" }, 400);
    let ultimo = "sin respuesta";
    for (const modelo of MODELOS) {
      const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent`, {
        method: "POST",
        headers: { "x-goog-api-key": Deno.env.get("GEMINI_API_KEY") ?? "", "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0, responseMimeType: "application/json" },
        }),
      });
      const j = await r.json().catch(() => ({}));
      if (r.ok) {
        const text = (j.candidates?.[0]?.content?.parts ?? []).map((p: { text?: string }) => p.text ?? "").join("");
        return json(h, { text, modelo, uso: j.usageMetadata });
      }
      ultimo = `${modelo}: ${j?.error?.message ?? r.status}`;
      if (![404, 429, 500, 503].includes(r.status)) break; // error de llave o de petición: no tiene caso probar otro modelo
    }
    return json(h, { error: ultimo }, 502);
  } catch (e) {
    return json(h, { error: String(e) }, 500);
  }
});
