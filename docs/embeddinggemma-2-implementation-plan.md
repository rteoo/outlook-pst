# EmbeddingGemma 2: local semantic mail search

Date: 2026-10-06. Status: proposed implementation; documentation only.

## Outcome and integration point

Find messages using paraphrases while retaining exact archive, folder, sender, recipient, subject, and date filters. The canonical CLI is `skills/outlook-pst/scripts/outlook_pst.py`; its current `matches` helper performs substring matching. Preserve existing commands, explicit mailbox selection, and default dry-run behavior.

## Milestones

1. **Corpus and baseline.** Define 40 held-out synthetic queries with known relevant messages, duplicate threads, missing bodies, and encoding differences. Measure current filtering/search coverage. Obtain explicit selection of any real archive used for a private evaluation; do not export it into the repository.
2. **Explicit index command.** Add proposed `semantic index` and `semantic search` subcommands, final syntax chosen to match the CLI. Index only caller-selected archive/folder scopes. Use a read-only archive copy for evidence work; no implicit Outlook attachment, mailbox change, export, or write-back.
3. **Message projection.** Index subject plus bounded body chunks with date and folder metadata. Preserve the backend's original locator and enough archive identity to detect replacement/reorganization. Do not assume libpff IDs or live Outlook EntryIDs survive every archive operation. Revalidate locators before opening a result.
4. **Runtime and storage.** Use a local text-only adapter with 768-dimensional normalized vectors, correct prefixes, timeouts, and explicit installed-model selection. Keep a rebuildable index outside the source archive and Git. Record archive identity, source hash, chunk version, model revision, and index completion state. Preserve installed-only operation; downloads are explicit.
5. **Retrieval.** Apply exact scope and metadata filters before semantic ranking. Add a bounded lexical candidate/ranking path if the baseline lacks one; do not call substring filtering BM25. Return subject, date, folder, snippet, score, and original reference in structured output. Runtime failure is actionable; users can still invoke existing list/show commands.
6. **Attachment extension.** After message retrieval passes, add explicitly selected attachment types with file-size/parser bounds and page citations. Avoid executing attachment content; keep extraction and source storage separate.

## Acceptance and verification

- Proposed gate: at least 85% Recall@5 across the held-out paraphrase set, and correct folder/date/sender exclusions for every fixture. Measure rather than promise production quality.
- Cover truncated indexes, archive changes, missing messages, mixed encodings, duplicate threads, cancellation, invalid vectors, and unavailable runtime.
- Test fake libpff/Outlook backends first. Live evidence requires an approved scratch archive and proof that its bytes remain unchanged.
- Run `python -B -W error::ResourceWarning -m unittest discover -s skills/outlook-pst/tests -v`; use existing Ruff if available. Check CLI help and plugin packaging. Preserve the builder's explicit file allowlist; do not bundle indexes, weights, or new modules accidentally.
- Record peak RSS, initial indexing time, and warm search latency. Every test removes temporary files and closes archives/processes even on failure.

## Rollout

Deliver message-only retrieval as opt-in CLI behavior. Preserve existing search commands for rollback. Attachment indexing and training are separate milestones. No dependencies, model files, private fixtures, or runtime data are added by this plan; their acquisition or modification needs appropriate authorization.

## References

- [Google model card](https://ai.google.dev/gemma/docs/embeddinggemma/model_card_2)
- [Unsloth runtime guide](https://unsloth.ai/docs/models/embeddinggemma-2)
