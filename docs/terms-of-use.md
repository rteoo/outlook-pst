# Outlook PST/OST terms of use

Effective date: 2026-10-29.

Publisher: RODRIGO TEODORO ALVES (GitHub: rteoo).
These terms describe use of the Outlook PST/OST plugin and its bundled local CLI.

## Software license

The software is provided under the [MIT License](../LICENSE). That license governs
copying, modification, distribution, attribution, warranty, and liability. These
terms do not replace or narrow the rights granted by that license.

## Intended use and requirements

Use the software with archives, mailboxes, and information you are authorized to
access. You choose the inputs, requested operations, and export destinations.
The plugin does not provide an email account or a hosted mailbox service.

Python and the relevant dependencies are required. Live Outlook access requires
Windows, Classic Outlook, and pywin32. Offline archive access uses libpff. AI
assistants and other third-party services have their own terms and requirements.

## Changes and exports

Review previews before authorizing changes. Applied changes can synchronize
through your email provider, and multi-item changes are not transactional. Keep
suitable backups before changing important data. Attaching a PST can modify the
archive. Use the pff backend for strictly offline reads.

Exports reconstruct content. Hashes establish exported-file integrity, not the
authenticity of original messages. Preserve originals when needed. Opening
exported HTML or attachments in another application is a separate action.

## Price, support, and privacy

The plugin does not sell products or process payments. Third-party platforms or
services may impose their own charges. Support is available through
<https://github.com/rteoo/outlook-pst/issues>, without a guaranteed response time
or service-level commitment. Do not post private mail or credentials there.

See the [privacy policy](privacy-policy.md) for local processing, assistant-platform
processing, storage, and support-information practices. Use the same issue page
for questions about these terms, without disclosing confidential information.

## Updates

Changes will be published in this repository with a revised effective date.
Git history records previous versions. The MIT License remains the software's
license unless a future release explicitly states otherwise.
