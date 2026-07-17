@echo off
rem Launch BF6 Settings Manager from source (requires uv: https://docs.astral.sh/uv/)
cd /d "%~dp0"
uv run python main.py %*
