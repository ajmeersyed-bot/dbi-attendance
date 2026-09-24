export default async function handler(req, res) {
  const APPS_SCRIPT_URL =
    "https://script.google.com/macros/s/AKfycbxbE5_9evU5AEkFemBNBfRsGGs6_W2nPsIWFIuTZrB5xpB2LQKSHM0WhzGewpxV6N5RLA/exec";

  try {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(req.body)
    });

    const text = await response.text();

    res.status(response.status);

    res.setHeader("Content-Type", "application/json");

    res.send(text);

  } catch (error) {

    res.status(500).json({
      success: false,
      error: error.message
    });

  }
}
