# Agent E2E Baseline Template Guide

This guide defines how to reuse the baseline E2E suite and report format.

## 1) Suite File Structure

Use `tests/agent_e2e_suite.json` as baseline template.

Required top-level blocks:

- `测试集`: name, version, bot_id, description
- `默认配置`: test user/email and stream settings
- `评测维度`: D1-D5 definitions with weights and pass rules
- `验收门槛`: numeric gates
- `用例`: case list

## 2) Case Schema (minimum)

Each case should include:

- `用例ID`
- `意图分类`
- `二级意图`
- `测试场景`
- `对话类型` (`单轮` / `多轮`)
- `会话组ID` (required for multi-turn)
- `轮次`
- `用户输入`
- `预期结果`
- `评测维度`
- `是否通过`, `失败原因`, `修复状态`, `备注`

## 3) Multi-turn Rules (must-follow)

- Keep all turns in the same `会话组ID`.
- Ensure turn order is explicit by `轮次`.
- Never split one group across disconnected runs unless using the same created conversation.

## 4) Iteration Tracking Convention

Use the `迭代追踪` sheet for full-case tracking (not badcase-only):

- `基线`: first full run (`PASS`/`FAIL`)
- `第1轮修复`, `第2轮回归`: follow-up rounds
- `当前状态`:
  - `✅稳定`: always pass
  - `✅已解决`: failed before, then passed continuously
  - `⚠️回归`: passed before, now fails
  - `⚠️待稳定`: fixed once but not yet stable
  - `❌未解决`: still failing

## 5) Scope Control

Before running full suite:

- Use `AGENT_E2E_START_INDEX` + `AGENT_E2E_MAX_CASES` for targeted checks.
- Only run full suite after owner approval.
- For unstable cases, run targeted multi-round checks first.
