# Source release: 2026.09.30-r11

Published scope: application source, browser UI, desktop launchers, anonymous offline tests and a blank configuration template.

## Verification

- 217 offline Python tests passed.
- 8 JavaScript files/inline blocks passed syntax checks.
- Source startup with blank credentials returned HTTP 200 for the console.
- Backend r11 and HTML r11-ui03.4 version markers matched.
- Publication privacy checks found no configured live credentials, non-placeholder email addresses, personal workspace identifiers or common token/private-key patterns.
- No production mailbox, LLM, work-order request or real review-state mutation was performed during release verification.

## Intentionally excluded

Original Git history, credentials/config.yaml, real customer examples, business spreadsheets, archived attachments, images, logs, databases, review history, login sessions, OCR binaries/models, Python environments and build output.

## Deployment requirements

Use config.example.yaml to create local configuration. Set internal-mail domains and audit aliases for the deployment. Import reference tables locally. Install OCR/7-Zip/Edge dependencies where needed. This release does not provision production credentials or data.

The battery editor matches the current r11 implementation. A fixed dropdown of five battery classes is not part of this snapshot.
