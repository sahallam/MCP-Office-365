# GitHub Repository Rulesets Configuration

This document provides instructions for configuring GitHub rulesets to restrict repository updates to known contributors and maintain code quality.

## Overview

GitHub rulesets allow you to control how people interact with branches and tags in your repository. They provide granular control over:
- Who can push to protected branches
- Required code reviews and approvals
- Status checks that must pass
- Commit signing requirements
- And much more

## Recommended Ruleset Configuration

### 1. Main Branch Protection

Protect your `main` branch from direct pushes and require pull requests with reviews.

#### Steps to Configure:

1. Go to your repository on GitHub
2. Click **Settings** → **Rules** → **Rulesets**
3. Click **New ruleset** → **New branch ruleset**
4. Configure as follows:

**Ruleset Name:** `Main Branch Protection`

**Enforcement Status:** Active

**Target Branches:**
- **Include:** `main` (or use pattern `main` or `master`)

**Bypass List (Who can bypass these rules):**
- ☑️ Repository administrator
- ☑️ Maintain role (optional - for trusted maintainers)

**Rules:**

- ☑️ **Restrict deletions**
  - Prevents anyone (except bypass users) from deleting the main branch

- ☑️ **Require a pull request before merging**
  - Required approvals: `1` (minimum - increase to 2+ for higher security)
  - ☑️ Dismiss stale pull request approvals when new commits are pushed
  - ☑️ Require review from Code Owners (if you have a CODEOWNERS file)
  - ☐ Require approval of the most recent reviewable push

- ☑️ **Require status checks to pass**
  - ☑️ Require branches to be up to date before merging
  - Add specific checks: (if you have CI/CD)
    - `build` (if you have a build workflow)
    - `test` (if you have tests)
    - `lint` (if you have linting)

- ☑️ **Block force pushes**
  - Prevents rewriting history on the main branch

- ☑️ **Require signed commits** (Recommended for security)
  - Requires all commits to be signed with GPG/SSH

- ☑️ **Require linear history** (Optional - prevents merge commits)
  - Enforces a clean, linear git history

### 2. Release Branch Protection

Protect release branches from unauthorized changes.

**Ruleset Name:** `Release Branch Protection`

**Target Branches:**
- **Include pattern:** `release/*` or `v*`

**Rules:**
- ☑️ Restrict deletions
- ☑️ Require a pull request before merging (2 required approvals)
- ☑️ Block force pushes
- ☑️ Require signed commits

### 3. Development Branch Rules

Allow more flexibility for development branches while maintaining some standards.

**Ruleset Name:** `Development Branch Standards`

**Target Branches:**
- **Include pattern:** `feature/*`, `bugfix/*`, `hotfix/*`

**Rules:**
- ☑️ Require signed commits (optional but recommended)
- ☐ Allow force pushes (to allow rebasing during development)

### 4. Tag Protection

Protect version tags from being deleted or modified.

**Ruleset Name:** `Tag Protection`

**Target Tags:**
- **Include pattern:** `v*` (matches v1.0.0, v2.1.3, etc.)

**Rules:**
- ☑️ Restrict deletions
- ☑️ Restrict updates (prevents tag moving)
- ☑️ Require signed commits

## Setting Up CODEOWNERS (Recommended)

Create a `.github/CODEOWNERS` file to automatically request reviews from specific people or teams:

```
# Default owners for everything in the repo
* @sahallam

# Authentication code requires security review
/src/auth.ts @sahallam

# Security-related files require careful review
/src/security.ts @sahallam
/src/audit.ts @sahallam

# Configuration files
/.env.example @sahallam
/tsconfig.json @sahallam
/package.json @sahallam
```

## Quick Setup Guide for This Repository

### Minimal Protection (Start Here)

For a public repository with few contributors:

1. **Main Branch Only Protection:**
   - Target: `main`
   - Require pull request: ✅ (1 approval)
   - Block force pushes: ✅
   - Require status checks: ✅ (if you have CI)

### Moderate Protection (Recommended)

For active development with multiple contributors:

1. **Main Branch Protection:** (as above) +
   - Required approvals: 2
   - Dismiss stale reviews: ✅
   - Require signed commits: ✅
   - Require linear history: ✅

2. **Tag Protection:**
   - Target: `v*`
   - Restrict deletions: ✅
   - Restrict updates: ✅

### Maximum Protection (Enterprise)

For critical production code:

1. **Main Branch Protection:** (as moderate) +
   - Required approvals: 2+
   - Require Code Owners review: ✅
   - Require conversation resolution: ✅
   - Require status checks to be up to date: ✅

2. **Release Branch Protection:**
   - Target: `release/*`
   - Required approvals: 2+
   - All main branch rules apply

3. **Tag Protection:** (as moderate)

4. **CODEOWNERS file:** ✅

## Managing Contributors

### Adding Known Contributors

1. Go to **Settings** → **Collaborators**
2. Click **Add people**
3. Search for GitHub username
4. Select role:
   - **Read**: Can view and clone
   - **Triage**: Can manage issues and PRs
   - **Write**: Can push to repository (subject to rulesets)
   - **Maintain**: Can manage repository without sensitive actions
   - **Admin**: Full access including settings

### Contributor Workflow

With rulesets enabled, contributors must:

1. Fork the repository (for external contributors)
2. Create a feature branch
3. Make changes and commit
4. Open a pull request to `main`
5. Wait for review approval
6. Maintainer merges after approval

## Restricting Fork and Push Access

### For Public Repositories

Since anyone can fork a public repository, you can:

1. **Disable force pushes** to prevent history rewrites
2. **Require pull requests** so all changes are reviewed
3. **Require approvals** from known contributors
4. **Use branch protection** to prevent direct commits
5. **Enable GitHub Actions workflow approval** for first-time contributors

### For Private Repositories

You have full control:

1. **Only invite trusted collaborators**
2. **Set appropriate role permissions** (Write, Maintain, Admin)
3. **Enable rulesets** to enforce standards even for collaborators
4. **Require 2FA** for all contributors (Organization setting)

## GitHub Actions Security

Restrict workflow execution to prevent malicious PRs from running actions:

1. Go to **Settings** → **Actions** → **General**
2. Under **Fork pull request workflows**:
   - ☑️ **Require approval for first-time contributors**
   - ☑️ **Require approval for all outside collaborators**

## Monitoring and Auditing

### Enable Security Features

1. **Dependabot alerts**: Settings → Security → Dependabot alerts
2. **Code scanning**: Settings → Security → Code scanning
3. **Secret scanning**: Settings → Security → Secret scanning

### Audit Log

For organizations, enable audit logging:
- Organization Settings → Audit log
- Monitor who created/modified rulesets
- Track bypass events

## Example Ruleset Configuration Files

GitHub rulesets can be exported/imported as JSON. Here's an example:

```json
{
  "name": "Main Branch Protection",
  "target": "branch",
  "enforcement": "active",
  "conditions": {
    "ref_name": {
      "include": ["refs/heads/main"],
      "exclude": []
    }
  },
  "rules": [
    {
      "type": "deletion"
    },
    {
      "type": "required_linear_history"
    },
    {
      "type": "pull_request",
      "parameters": {
        "required_approving_review_count": 1,
        "dismiss_stale_reviews_on_push": true,
        "require_code_owner_review": false,
        "require_last_push_approval": false
      }
    },
    {
      "type": "required_signatures"
    }
  ],
  "bypass_actors": [
    {
      "actor_id": 5,
      "actor_type": "RepositoryRole",
      "bypass_mode": "always"
    }
  ]
}
```

## Troubleshooting

### "Cannot push to protected branch"

**Solution:** Create a pull request instead of pushing directly.

### "Required status check is failing"

**Solution:** Ensure all CI checks pass before merging. Fix failing tests/builds.

### "You need at least 1 approval"

**Solution:** Ask a repository maintainer to review your PR.

### "Commit signature verification failed"

**Solution:** Set up GPG or SSH commit signing:
```bash
# For GPG
git config --global commit.gpgsign true

# For SSH (GitHub now supports this)
git config --global gpg.format ssh
git config --global user.signingkey ~/.ssh/id_ed25519.pub
```

## Best Practices

1. **Start with minimal protection** and increase as needed
2. **Document your workflow** in CONTRIBUTING.md
3. **Use meaningful branch names** (`feature/`, `bugfix/`, `hotfix/`)
4. **Keep rulesets simple** - don't over-complicate
5. **Test rulesets** with a non-critical branch first
6. **Communicate changes** to your team when updating rulesets
7. **Use bypass sparingly** - bypassing should be exceptional
8. **Review the audit log** regularly for bypass events

## Resources

- [GitHub Rulesets Documentation](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/about-rulesets)
- [Available Rules for Rulesets](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/available-rules-for-rulesets)
- [Creating Rulesets](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/creating-rulesets-for-a-repository)
- [CODEOWNERS Syntax](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-code-owners)

## For This Repository

### Recommended Initial Configuration

For the MCP-Office-365 repository, I recommend:

**Ruleset 1: Main Branch Protection**
- Target: `main`
- Require PR with 1 approval
- Block force pushes
- Require signed commits (optional but recommended)
- Block deletions

**Ruleset 2: Tag Protection**
- Target: `v*`
- Block deletions
- Block updates

**Additional:**
- Create `.github/CODEOWNERS` with yourself as owner
- Enable Dependabot alerts
- Enable secret scanning

This provides solid protection while remaining manageable for a solo maintainer or small team.
