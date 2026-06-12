# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Single-file desktop Pomodoro timer (`pomodoro.html`) — zero dependencies, no build step. Open in any browser to use.

## How to Run

```
# Open directly in default browser (Windows)
start pomodoro.html

# Or from PowerShell
Start-Process "f:\Ai-autoWk\pomodoro.html"
```

No server, no build, no package manager needed.

## Architecture

All code lives in `pomodoro.html` in three sections: `<style>` (CSS), `<body>` (HTML structure), `<script>` (JS logic).

### State Model

All app state is stored in a single `state` object, persisted to `localStorage` under key `pomodoro_app_data` via `StorageManager`:

```
state.settings      — workDuration, shortBreak, longBreak, longBreakInterval, autoStart, muted
state.timerState    — mode (work|shortBreak|longBreak), remainingSeconds, isRunning, pomodorosCompleted, currentTaskId
state.tasks[]       — { id, title, estimatedPomodoros, completedPomodoros, done, createdAt }
state.stats         — { daily: { "YYYY-MM-DD": count }, totalPomodoros, sessions[] }
```

On page refresh, the timer **always resets to paused** — running state is not restored between sessions.

### Timer Engine (drift-resistant)

Instead of `setInterval` counting down by 1 each second, the timer stores `_endAt = Date.now() + remainingSeconds * 1000` on start. Each `tick()` (fired every 200ms) computes `ceil((_endAt - Date.now()) / 1000)`. This prevents drift from timer jitter.

### AudioEngine

Creates sounds via Web Audio API oscillators (no audio files). Lazy-init on first user interaction (browser autoplay policy). Includes an `_ctxLimit` counter (max 30 concurrent oscillators) to prevent leaks.

### DOM Caching

DOM references accessed via `dom.*` getters (lazy, always current even if elements recreated) and `$()` helper for direct `getElementById`.

### Settings Factory

`makeSettingsHandler(key, min, max, defaultVal)` returns a change handler that validates, clamps, updates state, and persists — avoiding per-input boilerplate.

## Key Patterns

- `persist()` writes the entire `state` object to localStorage after every mutation
- `renderAll()` / individual `renderX()` functions are idempotent — they read state and rebuild DOM
- Task rendering uses `escapeHtml()` via a throwaway `<div>` to prevent XSS
- The progress ring uses SVG `stroke-dashoffset` calculated as `RING_CIRCUMFERENCE * (1 - remaining/total)`

## party_wechat_check.py — 党员微信群成员核对工具

Single-file Python utility to compare an Excel party-member roster against WeChat group member screenshots (OCR text). Finds who hasn't joined the group or hasn't changed their nickname.

### How to Run

```bash
# Install deps (once)
pip install rapidfuzz rich

# Basic usage
python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt"

# Interactive mode
python party_wechat_check.py

# Export results to Excel
python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt" --output "结果.xlsx"

# Show all match details
python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt" --verbose
```

### Workflow

1. User exports WeChat group member list via screenshot → WeChat built-in OCR "提取文字" → paste to `.txt`
2. Script reads Excel (auto-detects name column), parses WeChat text, runs two-pass fuzzy matching
3. Outputs: matched list, not-found list (need followup), uncertain list (manual verification needed)

### Key Design

- **Name normalization**: strips emoji, brackets, common WeChat decoration tags, compound surname awareness
- **Two-pass matching**: forward (Excel→WeChat contains-match) + reverse (extract Chinese name substrings from unmatched WeChat nicknames)
- **rapidfuzz**: `partial_ratio` for fuzzy matching when direct contains fails (default threshold 85%)
- **Rich**: colored console tables (auto-falls back to plain text if Rich not installed)
- **Excel output**: 4 sheets (核对结果, 未找到, 待确认, 微信群全量)
