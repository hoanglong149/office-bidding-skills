---
date: 2026-08-23
description: Index bộ skill diagram-design (39 biểu đồ editorial) đã port sang Hermes + theme IPC, lưu trong vault và ~/.hermes/skills.
tags: [skill, diagram, epc, tooling, ipc-brand]
type: reference
---

# Diagram Design — Skill biểu đồ editorial (39 loại)

**Nguồn:** `cathrynlavery/diagram-design` (MIT 2.6) — port về Hermes 2026-08-23.
**Vị trí skill (Hermes dùng):** `~/.hermes/skills/diagram-design/`
**Vị trí vault (ref + backup):** `40 Resources/Internal/DEV/SKILLs/diagram-design/`

## Mục đích

Sinh biểu đồ editorial sạch (HTML/SVG tự chứa, không Mermaid, không Figma) phục vụ báo cáo thầu EPC: sơ đồ công nghệ, tiến độ Gantt, quy trình thi công, cơ cấu tổ chức, luân chuyển vật liệu...

## Brand IPC (đã theme sẵn)

| Token | Màu |
|-------|-----|
| accent | `#1f6a99` (xanh dương IPC) |
| ink | `#16314d` (navy) |
| muted | `#5c7c8e` |
| paper | `#ffffff` |

Font: tiêu đề **Times New Roman**, thân/labels **Arial** (không cần mạng).

## Cách dùng 

1. Người dùng bảo: "vẽ sơ đồ [loại] về [nội dung]".
2. Em load skill, đọc `references/type-<loại>.md`.
3. Sinh HTML (copy từ `templates-epc/` hoặc `assets/example-<loại>.html` rồi điền data).
4. Render Chrome headless → PNG.
5. Gửi Telegram / ghép DOCX báo cáo thầu.

## 5 template EPC sẵn (`templates-epc/`)

- `template-architecture.html` — sơ đồ hệ thống xử lý khí thải, mặt bằng nhà máy
- `template-process.html` — quy trình thi công / vận hành
- `template-gantt.html` — tiến độ thi công
- `template-org-chart.html` — cơ cấu tổ chức dự án
- `template-data-flow.html` — sơ đồ công nghệ / luân chuyển vật liệu

## Cấu trúc thư mục

```
diagram-design/
├── SKILL.md              # chuẩn hóa: frontmatter + hướng dẫn EPC tiếng Việt
├── references/           # 53 file type-*.md (chi tiết từng loại)
├── assets/               # 142 example HTML (39 loại × 3 biến thể)
├── templates-epc/         # 5 template EPC đã theme IPC
├── scripts/              # build/lint/render scripts
├── README-EPC.md         # hướng dẫn ngắn
└── full-repo/            # toàn bộ repo gốc (prompts, commands, vendor) làm ref
```

## 39 loại biểu đồ

architecture, bar, bubble, data-flow, datalake, db-schema, dependency, deployment, dp-integration, dp-security-matrix, er, fishbone, flowchart, gantt, high-level, high-level-vertical, it-state, journey, kanban, layers, line, loop, medallion, nested, org-chart, polar, process, pyramid, quadrant, radar, ridgeline, sankey, scatter, sequence, slopegraph, state, story-map, swimlane, timeline, tree, treemap, uml-class, venn, wardley

## Lưu ý

- Target density 4/10 — trên 9 node thì tách 2 sơ đồ.
- Accent chỉ dành 1-2 node quan trọng nhất.
- Static HTML mặc định (không JS).
- Không dùng cho: sơ đồ unicode nhanh, bảng/danh sách đơn giản, "một hình" vô nghĩa.
