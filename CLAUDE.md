# Operations Manager: notes for Claude Code

An Obsidian plugin (id `operations-manager`) for tracking office equipment, services, contracts, licenses and bills in Markdown notes. The owner installs it with BRAT from `theshayneb/operations-manager`. Keep the plugin id: changing it would make it a different plugin.

## Code

- There is no build step: `main.js` is the source, plus `styles.css` and `manifest.json`.
- `OM_TEMPLATES` near the top of `main.js` lists the note types and their properties. The new-note dialog and the settings guide both read it, so update it when a type gains or loses a property.
- Don't make the plugin write to notes unless the owner asks for that action. Nothing should be written on load or while browsing. Per-note plugin state (alert switches, alert rules, history) goes in the plugin settings, keyed by file path.

## Workflow

- Push finished, checked work straight to `main`. No pull requests.
- Every change that ships gets a new patch version and a GitHub prerelease (BRAT installs prereleases):
  1. Bump `version` in `manifest.json`.
  2. Add a `## [x.y.z] - YYYY-MM-DD` entry at the top of `CHANGELOG.md`. The release workflow uses it as the release notes and fails without it.
  3. Commit and push to `main`, then run the "Release Plugin" workflow (`.github/workflows/release.yml`) on `main` with the input `version: x.y.z`. Pushing a tag `x.y.z` also works where tag pushes are allowed.
  4. Tell the owner the version number. Don't give them git commands; they use the GitHub web interface.
- Never reuse a version number that was already tagged or released. Note that tag `0.4.9` exists and is older than `0.4.10`.

## Checks before pushing

- `node --check main.js`
- Load the plugin against a stubbed `obsidian` module and render each tab and the settings tab, if the change touches rendering.
