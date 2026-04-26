# Repository Setup Instructions for Contributors

Complete these steps after merging the PR to enable the selective contributions model.

---

## 🔗 Step 1: Create the Pull Request (Do This First)

### Option A: Via Web Interface (Recommended)

1. **Go to the PR creation URL:**
   ```
   https://github.com/sahallam/MCP-Office-365/pull/new/claude/prepare-public-release-eQTFH
   ```

2. **Fill in the PR details:**
   - **Title:** `Prepare repository for v1.0.0 public release`
   - **Description:** Copy content from `PR_DESCRIPTION.md` (created in this repo)
   - **Base branch:** `claude/mcp-office365-main` (or `main` if that's your default)
   - **Compare branch:** `claude/prepare-public-release-eQTFH`

3. **Create Pull Request**

4. **Review and Merge** (no approval needed - you're the maintainer!)

### Option B: Via GitHub CLI (If Available)

```bash
gh pr create \
  --title "Prepare repository for v1.0.0 public release" \
  --body-file PR_DESCRIPTION.md \
  --base claude/mcp-office365-main \
  --head claude/prepare-public-release-eQTFH
```

---

## 🏷️ Step 2: Set Up Issue Labels (5 minutes)

Labels help organize issues and PRs during monthly reviews.

### Navigate to Labels

Go to: `https://github.com/sahallam/MCP-Office-365/labels`

### Create These Labels

| Label | Color | Description |
|-------|-------|-------------|
| `bug` | `#d73a4a` | Something isn't working |
| `documentation` | `#0075ca` | Improvements to documentation |
| `security` | `#ee0701` | Security vulnerability or fix |
| `needs-triage` | `#fbca04` | Needs initial review |
| `needs-info` | `#d876e3` | More information needed from reporter |
| `help-wanted` | `#008672` | Extra attention is needed |
| `wontfix` | `#ffffff` | This will not be worked on |
| `duplicate` | `#cfd3d7` | This issue or PR already exists |
| `good-first-issue` | `#7057ff` | Good for new contributors |

**How to create:**
1. Click "New label"
2. Enter name, description, and color
3. Click "Create label"
4. Repeat for each label

**Note:** Some labels like `bug` may already exist - that's fine!

---

## 📅 Step 3: Set Up Monthly Review Reminder (2 minutes)

Add a recurring calendar event to remind yourself to review issues/PRs.

### Calendar Event Details

```
Event: Review MCP-Office-365 Issues & PRs
Date: 1st of every month
Time: [Your preferred time]
Duration: 2 hours
Frequency: Monthly
Reminder: 1 day before

Description:
Monthly review of issues and pull requests for MCP-Office-365

Tasks:
1. Triage new issues (30 min)
2. Review pull requests (60 min)
3. Respond to comments (15 min)
4. Close duplicates/resolved issues (15 min)

Links:
- Issues: https://github.com/sahallam/MCP-Office-365/issues
- PRs: https://github.com/sahallam/MCP-Office-365/pulls
- Guide: https://github.com/sahallam/MCP-Office-365/blob/main/CONTRIBUTING.md
```

---

## 🔖 Step 4: Create v1.0.0 Release (10 minutes)

After merging the PR, create the official v1.0.0 release.

### A. Create Git Tag

```bash
# Switch to your default branch
git checkout claude/mcp-office365-main  # or 'main'

# Pull latest changes (including merged PR)
git pull

# Create annotated tag
git tag -a v1.0.0 -m "Release v1.0.0 - Initial public release"

# Push tag to GitHub
git push origin v1.0.0
```

### B. Create GitHub Release

1. **Go to:** https://github.com/sahallam/MCP-Office-365/releases/new

2. **Fill in release details:**
   - **Tag:** Select `v1.0.0` (the tag you just created)
   - **Release title:** `v1.0.0 - Initial Public Release`
   - **Description:** Copy from `RELEASE_NOTES.md` in your repository

3. **Options:**
   - ☑️ **Set as the latest release**
   - ☐ Set as a pre-release (leave unchecked)

4. **Click "Publish release"**

---

## 🛡️ Step 5: Enable Security Features (5 minutes)

### A. Dependabot Alerts

1. Go to: `https://github.com/sahallam/MCP-Office-365/settings/security_analysis`
2. Under "Dependabot alerts":
   - Click **Enable** (if not already enabled)
3. Under "Dependabot security updates":
   - Click **Enable** (optional - auto-creates PRs for security updates)

### B. Secret Scanning

1. Same page as above
2. Under "Secret scanning":
   - Click **Enable**
   - This detects accidentally committed secrets

### C. Code Scanning (Optional)

If you want automated code analysis:
1. Go to: `https://github.com/sahallam/MCP-Office-365/security/code-scanning`
2. Click **Set up code scanning**
3. Choose "CodeQL Analysis" → **Set up this workflow**
4. Commit the workflow file

**Note:** Code scanning requires GitHub Actions - optional for this project

---

## 🔐 Step 6: Configure GitHub Actions Security (2 minutes)

Prevent malicious PRs from running arbitrary code in GitHub Actions.

1. **Go to:** `https://github.com/sahallam/MCP-Office-365/settings/actions`

2. **Under "Fork pull request workflows":**
   - ☑️ **Require approval for first-time contributors**
   - ☑️ **Require approval for all outside collaborators**

3. **Click "Save"**

This ensures external PR workflows need your approval before running.

---

## 🛡️ Step 7: Set Up Branch Protection (Optional - 10 minutes)

Branch protection prevents accidental direct commits to main. See `GITHUB_RULESETS.md` for complete guide.

### Quick Setup (Minimal Protection)

1. **Go to:** `https://github.com/sahallam/MCP-Office-365/settings/rules`

2. **Click:** "New ruleset" → "New branch ruleset"

3. **Configure:**
   - **Name:** `Main Branch Protection`
   - **Status:** Active
   - **Target:** `main` (or your default branch name)

4. **Rules to enable:**
   - ☑️ **Restrict deletions**
   - ☑️ **Require a pull request before merging**
     - Required approvals: `1`
     - ☑️ Dismiss stale reviews on new commits
   - ☑️ **Block force pushes**

5. **Bypass list:**
   - ☑️ Repository administrators (so you can still push if needed)

6. **Click "Create"**

**Result:** All changes to main now require a PR (even from you). You can bypass if needed.

---

## 📧 Step 8: Prepare for First Contributors (2 minutes)

### A. Add Email to GitHub Profile (For Security Reports)

Security reports need a contact email. Ensure your GitHub profile has:
1. Go to: https://github.com/settings/profile
2. Add public email (or keep private and use GitHub's email)
3. This is how people will report security issues

### B. Verify CODEOWNERS Works

After merging:
1. Create a test PR (edit README.md)
2. Verify you're automatically requested as reviewer
3. Close the test PR

---

## 📊 Step 9: First Monthly Review (Scheduled)

On the 1st of next month, follow this workflow:

### 1. Review New Issues (30 min)

**Go to:** https://github.com/sahallam/MCP-Office-365/issues

**For each issue:**

```
✓ Valid bug?
  → Add label "bug"
  → Comment: "Thanks for reporting! I'll investigate this."
  
✓ Documentation issue?
  → Add label "documentation"
  → Quick fix if easy, or comment: "Will address in next release"
  
✓ Feature request?
  → Comment: "Thanks! This project accepts bug fixes only. For features, 
     please fork the repository. See CONTRIBUTING.md for details."
  → Add label "wontfix"
  → Close
  
✓ Duplicate?
  → Comment: "Duplicate of #[number]"
  → Add label "duplicate"
  → Close
  
✓ Need more info?
  → Add label "needs-info"
  → Comment asking for specific details
```

### 2. Review Pull Requests (60 min)

**Go to:** https://github.com/sahallam/MCP-Office-365/pulls

**For each PR:**

```
✓ Bug fix PR?
  → Review code changes
  → Check if issue is linked
  → Verify tests/reproduction steps
  → Test locally if critical
  → Merge or request changes
  
✓ Documentation PR?
  → Quick review
  → Merge if correct
  
✓ Feature PR?
  → Comment: "Thank you! This project accepts bug fixes only. For features,
     please maintain this in your fork. See CONTRIBUTING.md."
  → Close (don't merge)
  
✓ Security PR?
  → Priority review
  → Test thoroughly
  → Merge quickly
```

### 3. Clean Up (15 min)

- Close issues marked "wontfix" or "duplicate"
- Close stale "needs-info" issues (> 30 days no response)
- Update any documentation if needed

### 4. Done!

Close GitHub. See you next month! 🎉

---

## 🔄 Step 10: Adjust as Needed

### If You Get Too Many Issues/PRs

**Option 1: Extend Review Period**
- Change to quarterly (every 3 months)
- Update README and CONTRIBUTING.md with new timeline

**Option 2: Limit Scope Further**
- Only accept security fixes
- Close all other issues/PRs with fork guidance

**Option 3: Recruit Co-Maintainer**
- Find someone to help with reviews
- Add as collaborator with "Maintain" or "Triage" role
- Share the monthly review burden

### If It's Going Well

**Option 1: Review More Frequently**
- Switch to bi-weekly or weekly reviews
- Update documentation

**Option 2: Accept More Types**
- Consider accepting enhancement PRs
- Update CONTRIBUTING.md

**Option 3: Build Community**
- Enable GitHub Discussions
- Let users help each other
- Pin common issues as FAQ

---

## ✅ Setup Complete Checklist

After completing all steps, you should have:

- [ ] PR created and merged
- [ ] v1.0.0 tag created and pushed
- [ ] GitHub Release published
- [ ] Issue labels configured
- [ ] Monthly calendar reminder set
- [ ] Dependabot alerts enabled
- [ ] Secret scanning enabled
- [ ] GitHub Actions security configured
- [ ] Branch protection enabled (optional)
- [ ] First monthly review scheduled
- [ ] Email available for security reports

---

## 🎉 You're Ready!

Your repository is now set up for selective contributions with:

✅ Professional documentation and policies  
✅ Clear contribution guidelines  
✅ Structured issue/PR templates  
✅ Security features enabled  
✅ Sustainable monthly review process  
✅ Low maintenance burden (~2 hours/month)  

**Next milestone:** Wait for your first contributor! 🚀

---

## 📚 Quick Reference Links

**Repository:**
- Main: https://github.com/sahallam/MCP-Office-365
- Issues: https://github.com/sahallam/MCP-Office-365/issues
- PRs: https://github.com/sahallam/MCP-Office-365/pulls
- Settings: https://github.com/sahallam/MCP-Office-365/settings

**Documentation:**
- README: https://github.com/sahallam/MCP-Office-365/blob/main/README.md
- Contributing: https://github.com/sahallam/MCP-Office-365/blob/main/CONTRIBUTING.md
- Security: https://github.com/sahallam/MCP-Office-365/blob/main/.github/SECURITY.md
- Release Notes: https://github.com/sahallam/MCP-Office-365/blob/main/RELEASE_NOTES.md

**Monthly Review Workflow:**
1. Issues → Triage → Label → Comment
2. PRs → Review → Test → Merge or Close
3. Clean up stale items
4. Done until next month!

---

**Questions?** Review CONTRIBUTING.md or adjust the policy to fit your needs!
