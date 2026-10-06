# Pre-publication review — 2026-10-06

## Result

The standalone source is suitable for a public early release with the runtime
limitations below. The review reproduced and repaired reachable defects, added
regression coverage, and checked the complete publication tree for private data.
This is a manual code and release review, not a certification of native libpff,
Outlook, archive integrity, or absence of every possible defect.

## Scope and method

Reviewed the full mail CLI, parser, filters, file and COM backends, RTF decoder,
MIME reconstruction, export paths, mutation previews, package builder, regression
suite, manifests, documentation, and CI configuration in separate passes for
correctness, privacy/security, and portability. Findings were reproduced against
the copied source before fixes where practical. Tests use fictional mail and
reserved example domains. No real mailbox or Outlook profile was accessed.

The public repository starts with fresh history. It contains no imported private
repository commits, mail files, exports, attachments, credentials, machine paths,
or personal commit email. Only explicit source files were transferred.

## Findings and corrections

| ID | Severity | Reachable behavior before correction | Correction and proof |
| --- | --- | --- | --- |
| R1 | High when CSV is opened in a spreadsheet | Untrusted subject/sender values could begin with spreadsheet formulas in listings and manifests. | Neutralize common formula prefixes in CSV presentation fields. Regression tests cover listings, manifests, whitespace prefixes, and preservation of original message text. JSON and evidence files are not rewritten. |
| R2 | Medium | `--limit 0` acted as unlimited selection; negative limits selected an item. Reversed date ranges and empty replacement needles were accepted. | Reject these inputs before opening a file or Outlook namespace. Tests cover read and applied mutation entrypoints. |
| R3 | Medium | Successful libpff reads had no explicit close on success or consumer failure. | Close successfully opened handles in `finally`. Tests prove cleanup on both paths. Native libpff rejects closing an unopened handle, so failed opens retain their original error. |
| R4 | Medium | Live collections could contain contacts or appointments; explicit IDs could repeat the same item. | Process MailItems only, reject non-mail explicit IDs and case-insensitive duplicate IDs, and count only mail in the COM tree. Tests cover iteration, counts, and refusal before mutation. |
| R5 | Medium | RTF surrogate pairs could leave invalid Unicode and fail JSON/export encoding for emoji. | Combine UTF-16 pairs and visibly replace unpaired surrogates. Regression tests cover both cases. |
| R6 | Low | Deleted Items guard incorrectly treated similarly named sibling folders as descendants. | Compare the exact path and separator boundary, retaining refusal for Deleted Items and its children. Tests prove both behaviors and that Delete is never called on protected items. |
| R7 | Low | Warning state carried into later in-process CLI invocations. | Reset state at each invocation; regression test proves independent results. |
| R8 | Low | Invalid attachment extensions could reach Outlook profile calls; output validation could attempt to iterate a file as a directory. | Reject non-PST attachment targets and file export destinations before backend access. Export files and manifests use exclusive creation; filename selection also accounts for dangling symlinks. |
| R9 | Release integrity | The initial builder embedded duplicated manifest metadata and omitted license and release safety tests from the package. COM fake tests depended on native pywin32 and were skipped off Windows. | Read canonical root manifests, validate overlay agreement and listing length, include MIT license and safety tests, and use pure Python COM error fakes on every OS. Packaging tests cover relocation, explicit inclusion, exclusion of unlisted private artifacts, deterministic output, and refusal to overwrite existing output. |

`scripts/outlook_pst.py` and its tests live under `skills/outlook-pst/`.
Package generation lives in `skills/outlook-pst/scripts/build_plugin.py`.
Mailbox mutation remains opt-in through `--apply`; dry-run and applied-path tests
verify that callbacks are invoked only as intended.

## Verification

- Python 3.14 on Windows: all 41 regression tests passed, with no skipped tests.
- Ruff: selected correctness, import, and bugbear checks passed using the local
  `ruff.toml`; no tool or project dependency was installed for the review.
- Relocated plugin: CLI help and the bundled regression tests passed from a
  temporary directory independent of the checkout. Test directories clean up
  on failure as well as success.
- Native libpff 20260926, already cached locally: import and invalid synthetic
  archive rejection passed. This does not prove successful real-archive reads.
- Root portable manifest validated against the published Agent Plugins 1.0
  JSON schema. ZIP integrity and root/overlay metadata agreement were checked.
- CI is configured for Windows, Linux, and macOS on Python 3.10 and 3.14, using
  read-only workflow permissions and verified immutable official action commits.
  Check the public Actions run for live results; configuration alone is not proof.

## Remaining limits

Successful native archive reads, live Outlook mutations, timestamp behavior
against a real store, and plugin discovery in the desktop app have not been
verified for this release. Native dependency implementation and malformed-file
fuzzing are outside the completed review. Use approved scratch archives for
further runtime work; never substitute a real mailbox or evidence original.

COM scans read client-side properties and can be slow on very large mailboxes.
MIME assembly holds one message's attachments in memory. RTF conversion is a
focused text/HTML extractor, not a complete RTF interpreter. HTML and attachment
execution in other applications remain outside the CLI's boundary.

Selections are reevaluated between preview and `--apply`; specify explicit IDs
when binding an approval to items. Multi-item mutations are not transactional:
a later Outlook error can leave earlier changes applied. The CLI does not
provide a mailbox rollback. Outlook attachment can modify a PST. Exports rebuild
MIME and do not replace preservation of the original archive or authenticate it.

Keep export directories private and controlled. Concurrent hostile replacement
of their directory tree is outside the supported threat model. CSV protection
covers common formula prefixes; spreadsheet import behavior varies by client.

## Primary references

- [Outlook object classes](https://learn.microsoft.com/en-us/office/vba/api/outlook.olobjectclass)
- [MailItem.Delete behavior](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.delete)
- [Agent Plugins manifest schema](https://agent-plugins.org/schemas/1.0.0/plugin.schema.json)
- [OpenAI plugin packaging and local installation](https://developers.openai.com/plugins/build/plugins)
- [libpff-python distribution](https://pypi.org/project/libpff-python/)
- [pywin32 distribution](https://pypi.org/project/pywin32/)
