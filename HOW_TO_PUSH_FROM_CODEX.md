# Publishing Your Changes to GitHub from Codex

1. **Configure the remote**
   ```bash
   git remote -v              # inspect existing remotes
   git remote add origin https://github.com/<username>/<repo>.git
   ```
   Replace `<username>` and `<repo>` with your GitHub account and repository name. If the remote already exists, you only need to ensure it points to the correct URL.

2. **Authenticate (if prompted)**
   When you push, Git will request credentials. Use a GitHub personal access token in place of a password. You can create one at https://github.com/settings/tokens (classic) with the `repo` scope.

3. **Stage and commit your work**
   ```bash
   git status -sb             # confirm which files changed
   git add <files-or-dirs>
   git commit -m "Describe your change"
   ```

4. **Push the branch**
   ```bash
   git push -u origin <branch>
   ```
   The `-u` flag links your local branch with the remote one so later pushes can use simply `git push`.

5. **Open a pull request**
   Navigate to your repository on GitHub. You should see a banner offering to open a PR for the newly pushed branch. Click it, review the diff, and submit the PR.

6. **Update the branch later** (optional)
   ```bash
   git pull --rebase          # grab remote updates without merge commits
   git push
   ```

Following these steps will mirror your local Codex workspace changes back to GitHub.
