# Public edition current status / 公开版当前状态

- 更新时间 / Updated: 2026-08-28 (Asia/Shanghai)
- 当前阶段 / Stage: privacy-safe GitHub release audit
- 本轮目标 / Goal: build a standalone portfolio and graduate-application edition from the validated codebase without copying private assets or history
- 总体状态 / Overall status: 本地发布门槛已通过 / local release gate passed; remote creation and upload are waiting for the user's hotspot authorization
- 当前是否存在公开 EXE / Public EXE: no; executables are intentionally excluded
- 推荐启动方式 / Recommended launch: `python run.py --mode audit`

## 本轮完成内容 / Completed in this iteration

- Created a brand-new export with a fresh source tree and no inherited Git history.
- Removed every private-workbook fixed profile and added one entirely fictional `sim_demo` profile.
- Replaced the internal default template reference with a new public demo template generated from scratch.
- Added three deterministic SIM workbooks covering single/multi-row headers, bilingual XYZ, multi-sheet selection, L1–L13, merged cells, extra columns, and generated placeholder images.
- Added a bilingual portfolio README, architecture notes, privacy/testing documentation, third-party notices, sanitized UI previews, public tests, and a repository audit tool.
- Adapted namespace serialization in the public copy so templates using explicit OpenXML prefixes remain valid.
- Detected one unsafe background-screen capture during screenshot production, deleted it before Git initialization, and replaced it with a text-only SVG containing only fictional values. The unsafe capture was never committed or uploaded.

## 本轮文件变更 / File changes

| Path | Change | Reason | User impact |
| --- | --- | --- | --- |
| `src/bom_converter/` | copied and sanitized | retain generic conversion/UI core while removing private profiles | public source remains runnable with synthetic fixtures |
| `assets/public_demo_template.xlsx` | created | replace non-public template | safe default output template |
| `samples/simulated/` | created | independent public regression data | examples and tests need no real workbook |
| `tests/test_public_features.py` | created | public feature and privacy regressions | repeatable local validation |
| `tools/audit_public_repo.py` | created | release-gate scanning | blocks unexpected data and binaries |
| `README.md`, `docs/`, `SECURITY.md` | created | portfolio, architecture, privacy and safe-reporting guidance | understandable public entry point |
| `THIRD_PARTY_NOTICES.md` | created | record direct build dependencies and licenses | license provenance is explicit |
| `.gitignore` | created | exclude local state, outputs, binaries and sensitive categories | reduces accidental commits |

## 验证与测试结果 / Verification

- Synthetic workbook generation: completed with fixed seed `20260827`.
- Visual workbook review: completed for every sheet; no clipped key headers or unreadable sections were observed.
- Source quick conversion: 3 rows and 1 image, no verification error.
- Confirmed bilingual XYZ conversion: 3 rows and 2 images, no verification error.
- Full public automated suite: 20/20 passed with `PYTHONDONTWRITEBYTECODE=1`.
- Repository privacy/metadata scan: passed with 0 violations across 44 files before this status update.
- XLSX package audit: 4 explicitly allowed files; 0 hidden sheets, comments, external relationships, custom XML parts, or macro parts.
- Image audit: 2 PNG previews have 0 text-metadata chunks; 1 SVG has 0 external resources.
- Credential, local-path, excluded-identifier, large-file, and unexpected-binary scans: 0 findings in the final scanned set.
- GUI screenshots: quick and result pages use the fictional preview state; the mapping matrix is a text-only SVG.
- Microsoft Excel desktop acceptance of the public synthetic outputs: not part of the current source-only GitHub upload gate and not claimed as completed here.
- Public EXE build/start: not performed; no executable will be uploaded.
- GitHub CLI: version `2.98.0` is installed locally. Its authenticated remote state has intentionally not been queried during the offline phase.
- Local release gate on 2026-08-28: 20/20 tests passed and the repository audit passed with 0 violations across 44 files before Git initialization.
- Fresh local Git history: `main` has one root commit containing the 44 audited files; the worktree is clean and no remote is configured.
- Network state: no GitHub repository lookup, creation, visibility change, Release, or upload has been attempted. The intended remote remains `Rong67888/offline-bom-converter` as Private.

## 当前使用方法 / Usage

1. Install Python 3.10 or newer with Tkinter support.
2. Create and activate a virtual environment.
3. Run `python -m pip install -e .`.
4. Start audit mode with `python run.py --mode audit`, or quick mode with `python run.py --mode quick`.
5. Use `assets/public_demo_template.xlsx` for the fictional samples, or select a template you are authorized to process.
6. Outputs go to the chosen local folder. Business warnings appear on the result page; the optional JSON report remains local and is ignored by Git.

## 当前交付物 / Current artifacts

- runnable Python source under `src/bom_converter/`;
- GUI entry points `run.py` and `run_legacy.py`;
- public demo template and three synthetic workbooks;
- sanitized UI previews;
- public tests and privacy audit tool;
- bilingual documentation and third-party notices.

## 已知问题与限制 / Known limitations

- Automatic recommendations for abbreviations or low-information headers still require user confirmation.
- Only `.xlsx` is supported; macros, encryption, OLE objects, and online translation are unsupported.
- The public edition contains no private fixed profiles, so unknown formats intentionally enter generic audit.
- Synthetic tests do not replace local acceptance with authorized real workbooks.
- No public EXE, no no-Python-machine test, and no public Release are included.
- Project ownership and an open-source license remain unresolved; therefore no `LICENSE` file is present.

## 下一步计划 / Next steps

- P0: wait for the user to reply `热点已打开，可以继续`, then re-run the final privacy gate and perform the GitHub login check, Private repository creation, and push in one short network window.
- P1: review the private remote's file list and visibility, then update this status with the repository URL and commit SHA.
- P2: after ownership and license confirmation, decide whether the repository may become public; executable release remains a separate decision.

## 需要用户确认 / User decisions still required

- Confirm ownership of the source, UI, fictional template, and generic business rules before making the repository public.
- Choose an open-source license or keep the project source-available/closed before granting reuse rights.
- Confirm separately whether any future executable may be distributed.
- Provide no token in chat; GitHub authentication must be completed in the local CLI/browser flow.

## 交接与恢复说明 / Handoff prompt

Continue from the privacy-safe public export repository. GitHub CLI 2.98.0 is installed, but no authenticated remote query, repository creation, visibility change, Release, or upload has been attempted. The offline gate passed with 20/20 tests and 0 privacy violations. Confirm the fresh local `main` commit and clean tracked-file audit. Before any GitHub or Internet operation, tell the user exactly `即将连接GitHub，请打开热点。` and wait for the exact reply `热点已打开，可以继续`. Then re-run the final audit, check authentication without printing tokens, create `Rong67888/offline-bom-converter` as Private, push `main`, and verify visibility and the first commit in one network window. Stop if an unexpected workbook, binary, report, rule file, local path, identifier, or credential appears. Do not create a Release or upload an EXE. Keep the remote Private until ownership and license are explicitly confirmed.
