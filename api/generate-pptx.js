
export default async function handler(req, res) {
  if (req.method === "GET") {
    return res.status(200).json({
      status: "ok",
      service: "cement-daily-pptx-v4",
      version: "4.1"
    });
  }

  if (req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({
      error: "Method not allowed"
    });
  }

  const token = process.env.API_TOKEN;

  if (!token) {
    return res.status(503).json({
      error: "API_TOKEN not configured"
    });
  }

  if (req.headers.authorization !== `Bearer ${token}`) {
    return res.status(401).json({
      error: "Unauthorized"
    });
  }

  try {
    const data = typeof req.body === "string"
      ? JSON.parse(req.body)
      : req.body;

    if (!data || !data.report_date || !data.project_name) {
      return res.status(400).json({
        error: "Missing report_date or project_name"
      });
    }

    return res.status(200).json({
      status: "success",
      message: "施工日報資料接收成功",
      report_date: data.report_date,
      project_name: data.project_name,
      worker_count: data.worker_count ?? 0
    });

  } catch (error) {
    return res.status(400).json({
      error: String(error.message || error)
    });
  }
}
