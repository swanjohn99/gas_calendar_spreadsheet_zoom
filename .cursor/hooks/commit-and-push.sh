#!/usr/bin/env bash
# Fallback commit+push after the agent turn ends (stop). Fail open.
# If the agent already committed, this no-ops.
set -u

# Drain stdin (hook JSON payload).
cat >/dev/null || true

ROOT="$(git rev-parse --show-toplevel 2>/dev/null || true)"
if [[ -z "${ROOT}" ]]; then
  exit 0
fi
cd "${ROOT}" || exit 0

exec 9>"${ROOT}/.git/auto-commit.lock"
flock -w 120 9 || exit 0

if git diff --quiet && git diff --cached --quiet && [[ -z "$(git ls-files --others --exclude-standard)" ]]; then
  exit 0
fi

git add -A -- . 2>/dev/null || true

# Drop staged secrets / credential-like paths (align with deploy-after-changes.mdc)
is_blocked_() {
  local f="$1"
  local base
  base="$(basename "$f")"
  case "$base" in
    .env|.env.*|.clasp.json|.clasprc.json|credentials.json|*.pem|*.key|debug-*.log)
      return 0
      ;;
  esac
  case "$f" in
    *.env|*/.env|*/.env.*|*credentials*|*secret*|.clasp.json|*/.clasp.json|.clasprc.json|*/.clasprc.json|.cursor/debug-*.log)
      return 0
      ;;
  esac
  return 1
}

# Unstage adds/modifies of blocked paths; keep deletions (removing secrets from git).
while IFS= read -r f; do
  [[ -z "$f" ]] && continue
  if is_blocked_ "$f"; then
    git reset -q -- "$f" 2>/dev/null || true
  fi
done < <(git diff --cached --diff-filter=ACMR --name-only 2>/dev/null)

if git diff --cached --quiet; then
  exit 0
fi

# Resolve author: env -> local git config -> last commit -> hard fallback.
_name="${GIT_AUTHOR_NAME:-}"
_email="${GIT_AUTHOR_EMAIL:-}"
[[ -z "$_name" ]] && _name="$(git config --local user.name 2>/dev/null || true)"
[[ -z "$_email" ]] && _email="$(git config --local user.email 2>/dev/null || true)"
[[ -z "$_name" ]] && _name="$(git config user.name 2>/dev/null || true)"
[[ -z "$_email" ]] && _email="$(git config user.email 2>/dev/null || true)"
[[ -z "$_name" ]] && _name="$(git log -1 --format='%an' 2>/dev/null || true)"
[[ -z "$_email" ]] && _email="$(git log -1 --format='%ae' 2>/dev/null || true)"
[[ -z "$_name" ]] && _name="swanjohn99"
[[ -z "$_email" ]] && _email="bora.india@gmail.com"

export GIT_AUTHOR_NAME="$_name"
export GIT_AUTHOR_EMAIL="$_email"
export GIT_COMMITTER_NAME="${GIT_COMMITTER_NAME:-$_name}"
export GIT_COMMITTER_EMAIL="${GIT_COMMITTER_EMAIL:-$_email}"

if ! git commit -m "auto: sync agent file changes" >/dev/null 2>&1; then
  exit 0
fi

# Concurrent hooks can win the GitHub CAS first; treat "already at HEAD" as success.
in_sync_() {
  local remote_head
  remote_head="$(git rev-parse -q --verify origin/main 2>/dev/null || true)"
  [[ -n "$remote_head" && "$remote_head" == "$(git rev-parse HEAD)" ]]
}

push_head_() {
  local i
  for i in 1 2 3; do
    git fetch origin --prune >/dev/null 2>&1 || true
    if in_sync_; then
      return 0
    fi
    if git rev-parse -q --verify origin/main >/dev/null \
        && ! git merge-base --is-ancestor origin/main HEAD 2>/dev/null; then
      git pull --rebase origin main >/dev/null 2>&1 || true
      if in_sync_; then
        return 0
      fi
    fi
    if git push -u origin HEAD >/dev/null 2>&1; then
      return 0
    fi
    sleep 1
  done
}

push_head_
exit 0
