# Contributing to spreadsheet-tools

Thank you for your interest. All contributions go through pull requests and are reviewed by the maintainer.

## Author

This project is maintained solely by **Keiner Alvarado**. All commits, releases, and authorship records must attribute **Keiner Alvarado** as the original author.

---

## Getting started

```bash
git clone https://github.com/keinerdeveloper/spreadsheet-tools.git
cd spreadsheet-tools
uv sync
```

## Development workflow

1. Fork the repository.
2. Create a branch from `main`:
   ```bash
   git checkout -b feat/your-feature
   ```
3. Make focused, minimal changes.
4. Verify all checks pass:
   ```bash
   uv run ruff check src/
   uv run ruff format --check src/
   uv run python -m pytest
   ```
5. Open a pull request against `main`.

## Code standards

- **Correctness first** — no silent failures, no swallowed exceptions.
- **Minimal diffs** — change only what the issue requires.
- **No decorative edits** — no style-only refactors unless the PR is explicitly about style.
- **Python ≥ 3.13** — use modern type syntax (`X | Y`, `list[str]`, etc.).
- **JSON output contract** — all CLI commands must return valid JSON to `stdout`; errors go to `stderr`.

## Commit messages

Use conventional commits:

```
feat: add --include-formulas flag to read-range
fix: handle merged cells in sheet-info
docs: update SKILL.md with copy-sheet examples
chore: bump openpyxl to 3.1.5
```

## Reporting bugs

Open an issue using the **Bug report** template. Include:
- Command run (sanitize any sensitive paths)
- Full `stderr` output
- Python version (`python --version`)
- `uv` version (`uv --version`)

## Requesting features

Open an issue using the **Feature request** template. Describe the use case clearly — not just the proposed API.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE) and that copyright remains with **Keiner Alvarado**.
