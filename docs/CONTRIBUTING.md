# Contributing to Office4AI

## Development Environment Setup

```bash
# Clone the repository
git clone https://github.com/JQQ/office4ai.git
cd office4ai

# Install all dependencies (including dev tools)
uv sync --all-extras
```

## Common Commands

The project uses [poethepoet](https://github.com/nat-n/poethepoet) as task runner:

```bash
# Code quality
poe lint              # Run ruff linter
poe lint-fix          # Auto-fix lint issues
poe format            # Format code with ruff
poe format-check      # Check code formatting
poe typecheck         # Run mypy type checking

# Testing
poe test              # Run all tests
poe test-unit         # Unit tests only
poe test-integration  # Integration tests only
poe test-cov          # Tests with coverage report
poe test-verbose      # Tests in verbose mode

# Combined tasks
poe check             # lint + format-check + typecheck
poe fix               # lint-fix + format
poe pre-commit        # format + lint-fix + typecheck + test

# Cleanup
poe clean             # Remove .pyc files and coverage artifacts
```

## Code Conventions

| Rule | Value |
|------|-------|
| Line length | 120 characters |
| Python version | 3.10+ (recommend 3.11) |
| Type annotations | Required (`mypy disallow_untyped_defs = true`) |
| Import ordering | stdlib > third-party > local (ruff auto-sorts) |

### Ruff Rules

```toml
select = ["E", "W", "F", "I", "B", "C4", "UP"]
ignore = ["E501", "B008", "C901"]
"__init__.py" = ["F401"]  # Allow unused imports in __init__
```

## Testing

Tests are organized under `tests/`:

```
tests/
├── conftest.py          # Shared fixtures
├── unit_tests/          # Unit tests (isolated, no external deps)
├── integration_tests/   # Integration tests (@pytest.mark.integration)
└── contract_tests/      # Contract tests (@pytest.mark.contract)
```

Run before submitting a PR:

```bash
poe pre-commit
```

## Versioning and Releases

Office4AI has two independent versions:

- The package version in `pyproject.toml` identifies an Office4AI software release.
- `OASP_PROTOCOL_VERSION` in `office4ai/environment/workspace/socketio/versioning.py` identifies the OASP wire contract used by the Socket.IO handshake.

Routine package PATCH or MINOR releases must not change `OASP_PROTOCOL_VERSION`. Update the protocol version only when the OASP
wire contract or compatibility boundary changes. A protocol-version change must be coordinated with the Office Add-In, covered by
the real Socket.IO contract tests, and recorded in the release notes.

## Project Structure

```
office4ai/
├── base.py                  # OfficeEnv base class
├── schema.py                # Data models (OfficeAction, OfficeObservation)
├── exceptions.py            # Exception definitions
├── certs/                   # Auto certificate management
├── dtos/                    # A2C protocol data transfer objects
│   ├── common.py            # Common types
│   ├── word.py              # Word DTOs (13 events)
│   ├── ppt.py               # PPT DTOs (10 events)
│   └── excel.py             # Excel DTOs (4 events)
├── environment/
│   └── workspace/
│       ├── office_workspace.py  # OfficeWorkspace (Socket.IO server lifecycle)
│       └── socketio/
│           ├── server.py        # Socket.IO server
│           └── namespaces/      # /word, /ppt, /excel namespaces
├── a2c_smcp/                # MCP Server infrastructure
│   ├── server.py            # BaseMCPServer
│   ├── config.py            # MCPServerConfig
│   ├── tools/               # MCP Tools
│   │   ├── base.py          # BaseTool (declarative pattern)
│   │   └── word/            # 9 Word tools (MVP)
│   └── resources/           # MCP Resources
└── office/
    └── mcp/
        └── server.py        # OfficeMCPServer (entry point)
```

## Contribution Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Make your changes
4. Run `poe pre-commit` to verify everything passes
5. Commit your changes (`git commit -m 'Add your feature'`)
6. Push to your branch (`git push origin feature/your-feature`)
7. Open a Pull Request against `main`
