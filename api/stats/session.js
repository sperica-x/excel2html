const { registerVisit } = require("./_store");

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const payload =
      typeof req.body === "string" && req.body
        ? JSON.parse(req.body)
        : req.body;
    const stats = await registerVisit(payload);
    return res.status(200).json(stats);
  } catch (error) {
    const statusCode = error.statusCode || 500;
    return res.status(statusCode).json({
      error: statusCode === 500 ? "Failed to register visit" : error.message,
    });
  }
};
