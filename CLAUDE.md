# sheet-call-tree

## Environment

- Use `uv` for all Python environment and package management (`uv` is at `/home/rok/.local/bin/uv`)
- Create/recreate the venv: `uv venv --clear`
- Install dependencies: `uv pip install -e ".[dev]"`
- Run commands in venv: `.venv/bin/python` or `.venv/bin/pytest`
- Do NOT use `pip`, `python3 -m pip`, or `apt` for package management
