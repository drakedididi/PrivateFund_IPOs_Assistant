const DATA_BASE_URL = "https://pub-726d3f620e2c4dbf99889a63c9a5687d.r2.dev";
const ALLOWED_FILES = new Set([
  "Asharecalendar_data.json",
  "bondcalendar_data.json",
  "Hsharecalendar_data.json",
  "calendar_data.json",
  "trading_holidays.json",
]);

export async function onRequestGet(context) {
  const fileName = context.params.file;
  if (!ALLOWED_FILES.has(fileName)) {
    return new Response("Not found", { status: 404 });
  }

  const upstream = await fetch(`${DATA_BASE_URL}/${fileName}`, {
    headers: { Accept: "application/json" },
    cf: { cacheTtl: 300, cacheEverything: true },
  });

  if (!upstream.ok) {
    return new Response("Calendar data unavailable", { status: upstream.status });
  }

  return new Response(upstream.body, {
    status: 200,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "public, max-age=300",
      "Access-Control-Allow-Origin": "*",
    },
  });
}
