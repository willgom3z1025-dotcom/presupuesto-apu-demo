// Supabase Edge Function: proponer-partida
// Recibe {prompt} desde la aplicación (GitHub Pages) y lo envía al modelo de IA.
// La llave de la API vive solo aquí, como secreto del proyecto (ANTHROPIC_API_KEY); nunca en el HTML.
// Despliegue:
//   supabase secrets set ANTHROPIC_API_KEY=...        (en el proyecto presupuestos-apu)
//   supabase functions deploy proponer-partida
// En la aplicación: Proyecto > Inteligencia artificial > URL = https://<ref>.supabase.co/functions/v1/proponer-partida
// y la anon key del proyecto (es pública por diseño).

const ORIGENES = ["https://willgom3z1025-dotcom.github.io"];
const MODELO = "claude-haiku-5-5";
const MAX_PROMPT = 60000; // caracteres; el catálogo estándar ocupa unos 20,000

function cors(origin: string | null) {
  const ok = origin && ORIGENES.includes(origin) ? origin : ORIGENES[0];
  return {
    "Access-Control-Allow-Origin": ok,
    "Access-Control-Allow-Headers": "authorization, apikey, content-type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  };
}

Deno.serve(async (req) => {
  const h = cors(req.headers.get("origin"));
  if (req.method === "OPTIONS") return new Response("ok", { headers: h });
  if (req.method !== "POST") return new Response("Método no permitido", { status: 405, headers: h });
  try {
    const { prompt } = await req.json();
    if (typeof prompt !== "string" || !prompt || prompt.length > MAX_PROMPT) {
      return new Response(JSON.stringify({ error: "prompt inválido" }), { status: 400, headers: { ...h, "Content-Type": "application/json" } });
    }
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": Deno.env.get("ANTHROPIC_API_KEY") ?? "",
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({ model: MODELO, max_tokens: 4000, messages: [{ role: "user", content: prompt }] }),
    });
    const j = await r.json();
    if (!r.ok) return new Response(JSON.stringify({ error: j?.error?.message ?? "error del proveedor" }), { status: 502, headers: { ...h, "Content-Type": "application/json" } });
    const text = (j.content ?? []).filter((c: { type: string }) => c.type === "text").map((c: { text: string }) => c.text).join("");
    return new Response(JSON.stringify({ text, modelo: MODELO, uso: j.usage }), { headers: { ...h, "Content-Type": "application/json" } });
  } catch (e) {
    return new Response(JSON.stringify({ error: String(e) }), { status: 500, headers: { ...h, "Content-Type": "application/json" } });
  }
});
