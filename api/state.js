// Vercel serverless function: GET /api/state + POST /api/state
// Reads/writes state.json via GitHub Contents API (token server-side)
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const GITHUB_REPO = 'Alinoerie/kitchen-roadmap';
const GITHUB_PATH = 'data/state.json';
const GITHUB_BRANCH = 'master';

const githubApi = async (path, method = 'GET', body = null) => {
  const url = `https://api.github.com/repos/${GITHUB_REPO}/contents/${path}`;
  const opts = {
    method,
    headers: {
      'Authorization': `token ${GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'kitchen-roadmap-app',
    },
  };
  if (body) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(url + `?ref=${GITHUB_BRANCH}`, opts);
  return res.json();
};

export default async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (!GITHUB_TOKEN) {
    return res.status(500).json({ error: 'GITHUB_TOKEN not configured' });
  }

  try {
    if (req.method === 'GET') {
      const data = await githubApi(GITHUB_PATH);
      if (data.content) {
        const content = Buffer.from(data.content, 'base64').toString('utf-8');
        const state = JSON.parse(content);
        return res.status(200).json(state);
      }
      return res.status(404).json({ error: 'state.json not found' });
    }

    if (req.method === 'POST') {
      const { state, savedBy } = req.body;
      if (!state) return res.status(400).json({ error: 'state required' });

      // Get current file SHA (required for update)
      const current = await githubApi(GITHUB_PATH);
      if (!current.sha) {
        return res.status(500).json({ error: 'Could not fetch current state SHA' });
      }

      // Prepare update
      const newContent = Buffer.from(JSON.stringify(state, null, 2)).toString('base64');
      const body = {
        message: `Update state: ${savedBy || 'anon'} @ ${new Date().toISOString()}`,
        content: newContent,
        sha: current.sha,
        branch: GITHUB_BRANCH,
      };

      const result = await fetch(
        `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_PATH}`,
        {
          method: 'PUT',
          headers: {
            'Authorization': `token ${GITHUB_TOKEN}`,
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'kitchen-roadmap-app',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(body),
        }
      );

      if (result.ok) {
        const resultData = await result.json();
        return res.status(200).json({
          success: true,
          commitSha: resultData.commit.sha,
          commitUrl: resultData.commit.html_url,
        });
      } else {
        const err = await result.json();
        return res.status(result.status).json({ error: err.message || 'GitHub API error' });
      }
    }

    return res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}