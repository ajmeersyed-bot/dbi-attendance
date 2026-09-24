export default async function handler(req, res) {
  const APPS_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbz0jb8yBsEJE1iQsZdDIS2-6JHj65Yft-rTxhvuvWLtRsc-rjA4pBzS0SloR0CWxhXJlQ/exec";

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
