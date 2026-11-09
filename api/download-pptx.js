
const OWNER = 'pikachiu19840222-glitch';
const REPO = 'cement-daily-pptx-v4';

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({
      error: 'Method not allowed'
    });
  }

  const apiToken = process.env.API_TOKEN;
  const githubToken = process.env.GITHUB_TOKEN;

  if (!apiToken || !githubToken) {
    return res.status(503).json({
      error: 'Server credentials not configured'
    });
  }

  if (
    req.headers.authorization !==
    `Bearer ${apiToken}`
  ) {
    return res.status(401).json({
      error: 'Unauthorized'
    });
  }

  const artifactId = String(
    req.query.artifact_id ?? ''
  );

  if (!/^\d+$/.test(artifactId)) {
    return res.status(400).json({
      error: 'Valid artifact_id is required'
    });
  }

  try {
    const url =
      `https://api.github.com/repos/${OWNER}/${REPO}` +
      `/actions/artifacts/${artifactId}/zip`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${githubToken}`,
        Accept: 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'cement-daily-pptx-v4'
      },
      redirect: 'follow'
    });

    if (!response.ok) {
      return res.status(
        response.status === 404 ? 404 : 502
      ).json({
        error: 'Unable to download GitHub artifact',
        upstream_status: response.status
      });
    }

    const zip = Buffer.from(
      await response.arrayBuffer()
    );

    if (
      zip.length < 4 ||
      zip.subarray(0, 2).toString() !== 'PK'
    ) {
      return res.status(502).json({
        error: 'GitHub did not return a ZIP archive'
      });
    }

    res.setHeader(
      'Content-Type',
      'application/zip'
    );

    res.setHeader(
      'Content-Disposition',
      `attachment; filename="cement-daily-report-${artifactId}.zip"`
    );

    res.setHeader('Cache-Control', 'no-store');

    return res.status(200).send(zip);

  } catch (error) {
    return res.status(502).json({
      error: 'Artifact download failed'
    });
  }
}
