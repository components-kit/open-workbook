# Agent Notes

- Do not run a full build after every routine code or visual change. Prefer focused checks such as TypeScript, lint, or targeted tests when they are relevant.
- Husky runs the project build on commit, so reserve manual builds for broader changes, release checks, or when a build failure is likely from the edit.
- Before every commit, recheck whether docs need updates. Review `README.md`, `docs/`, `CHANGELOG.md`, `llms.txt`, and generated `llms-full.txt` when behavior, commands, package metadata, tool surface, or release process changes.
