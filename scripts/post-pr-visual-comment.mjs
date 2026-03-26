const githubToken = process.env.GITHUB_TOKEN;
const repository = process.env.GITHUB_REPOSITORY;
const prNumber = process.env.PR_NUMBER;
const commentBody = process.env.PR_VISUAL_COMMENT_BODY;

if (!githubToken || !repository || !prNumber || !commentBody) {
  process.exit(0);
}

const [owner, repo] = repository.split("/");
const marker = "<!-- pptx-react-renderer-visual-pr-comment -->";

const commentsResponse = await githubApi(
  `https://api.github.com/repos/${owner}/${repo}/issues/${prNumber}/comments?per_page=100`,
  {
    headers: {
      Authorization: `Bearer ${githubToken}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
    },
  },
);
const comments = await commentsResponse.json();
const existing = comments.find((comment) => typeof comment.body === "string" && comment.body.includes(marker));

if (existing) {
  await githubApi(`https://api.github.com/repos/${owner}/${repo}/issues/comments/${existing.id}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${githubToken}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ body: commentBody }),
  });
} else {
  await githubApi(`https://api.github.com/repos/${owner}/${repo}/issues/${prNumber}/comments`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${githubToken}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ body: commentBody }),
  });
}

async function githubApi(url, options) {
  const response = await fetch(url, options);

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`GitHub API request failed (${response.status}): ${body}`);
  }

  return response;
}
