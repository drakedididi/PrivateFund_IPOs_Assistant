const ALLOWED_FILES = new Set([
  "Asharecalendar_data.json",
  "bondcalendar_data.json",
  "Hsharecalendar_data.json",
  "calendar_data.json",
  "trading_holidays.json",
]);
const R2_BINDING_NAMES = ["PRIVATEFUND_DATA", "CALENDAR_DATA", "R2_BUCKET"];

export async function onRequestGet(context) {
  const fileName = context.params.file;
  if (!ALLOWED_FILES.has(fileName)) {
    return new Response("Not found", { status: 404 });
  }

  const bucket = R2_BINDING_NAMES
    .map((name) => context.env[name])
    .find((binding) => binding && typeof binding.get === "function");

  if (!bucket) {
    return new Response("R2 binding is not configured", { status: 500 });
  }

  const object = await bucket.get(fileName);
  if (!object) {
    return new Response("Calendar data not found", { status: 404 });
  }

  return new Response(object.body, {
    status: 200,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "public, max-age=300",
      "Access-Control-Allow-Origin": "*",
    },
  });
}
