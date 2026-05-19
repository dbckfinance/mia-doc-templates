# mia-doc-templates

Cursor SDK workspace for M&IA document generation.

This repository is cloned by Cursor cloud agents when generating documents
(CIMs, teasers, pitchbooks, DCF models, LBO models, memos, reports).

## Structure

- `.cursor/` — MCP config, agent definitions, skills
- `generators/` — Node.js scripts for each document type (pptx, xlsx, docx, pdf)
- `shared/` — Brand constants and financial helper functions
- `public/` — Static assets (logo)

## Usage

This repo is not meant to be used directly. It serves as the workspace
cloned by Cursor SDK cloud agents via `CURSOR_SDK_DEFAULT_REPO`.

Generated documents are extracted via `harvestArtifact()` and uploaded
to Supabase Storage. Work branches are purged automatically every 24h.

## Dependencies

Pre-installed so `npm install` in the VM is a no-op:

- pptxgenjs — PowerPoint generation
- exceljs — Excel generation
- docx — Word document generation
- pdfkit — PDF generation
- jspdf — Alternative PDF generation
- lodash-es — Utility functions
