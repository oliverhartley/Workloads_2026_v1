---
name: appscript-setup-sync
description: Sets up a Google Apps Script project using clasp (via npx) and synchronizes it with a GitHub repository. Use this when the user wants to start a new GAS project or link an existing one to GitHub.
---

# Setup Appscript & GitHub Sync

This skill automates the connection between a local workspace, a Google Apps Script (GAS) project, and a GitHub repository using `npx @google/clasp`.

## Workflow Steps

### 1. Project Initialization
- Ask the user for the **Google Apps Script Project ID**.
- Check if a `.clasp.json` already exists. If not, run `npx @google/clasp clone <ProjectID>`.
- **Note:** If the user is not logged in, instruct them to run `npx @google/clasp login` in the terminal manually first (as it requires a browser).

### 2. GitHub Integration
- Ask the user if they want to sync to a GitHub repository.
- If yes, ask for the **Repository URL**.
- If the local folder is not a git repo, run `git init`.
- Add the remote: `git remote add origin <URL>`.
- Create a `.gitignore` if it doesn't exist, ensuring it includes `node_modules/` and potentially sensitive clasp files if necessary.

### 3. Continuous Sync Logic
After setup, you (the agent) are responsible for maintaining the "Triple Sync":
1. **Local to GAS:** After any code change, run `npx @google/clasp push`.
2. **Local to GitHub:** After a successful push to GAS, stage the changes, create a meaningful commit message based on the edits made, and run `git push origin main` (or the current branch).

## Constraints
- **NEVER** attempt to install `clasp` via `npm install -g`. Always use `npx @google/clasp`.
- Ensure `appsscript.json` is always included in the sync.
- If a sync conflict occurs, alert the user before overwriting remote GAS code.
