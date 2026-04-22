const { summarizeStats } = require("./_store");

module.exports = async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const stats = await summarizeStats();
    return res.status(200).json(stats);
  } catch (error) {
    return res.status(500).json({ error: "Failed to load statistics" });
  }
};
