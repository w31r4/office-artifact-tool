# Default Template Library

This checked-in catalog provides 20 independently selectable, reference-backed Office templates: 7 documents, 7 presentations, and 6 spreadsheets.

## Layout

```text
skills/default-template-library/
├── manifest.json
├── assets/icon.svg
└── skills/artifact-template-<name>/
    ├── SKILL.md
    ├── artifact-template.json
    ├── agents/agent.yaml
    └── assets/
        ├── preview.png
        └── reference.docx | reference.pptx | reference.xlsx
```

Each nested skill retains its reference Office file and preview image. Use the named template skill to create a new artifact while preserving the retained layout and formatting unless the request calls for a change.

These resources are repository-only and are intentionally excluded from the npm package tarball.
