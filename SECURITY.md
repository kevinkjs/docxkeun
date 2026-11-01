# Security

docxkeun opens documents from untrusted sources and implements its own cryptography, so security reports are very welcome.

## Reporting a vulnerability

Please **do not open a public issue**. Report it privately through GitHub: go to the repository's **Security** tab and choose **Report a vulnerability**. Include a minimal document or snippet that reproduces the problem if you can.

You can expect a first reply within a week. Fixes are released as a new version and credited in the changelog unless you prefer otherwise.

## Supported versions

While docxkeun is 0.x, only the latest release receives security fixes.

## What docxkeun does and does not protect

- **Untrusted documents.** Parsing never loads external entities or resources. Rendered text is escaped. Links open only with allowed URL schemes (`http:`, `https:`, `mailto:`, `tel:` by default). Images referenced by URL are not loaded unless the host opts in. Crashes, hangs or memory blow-ups caused by crafted files count as security issues.
- **Encryption is real.** Password-to-open `.docx` files (AES-256, MS-OFFCRYPTO Agile) and PDF user passwords (AES-256, revision 6) encrypt content. The implementations are docxkeun's own and **have not been independently audited**. They are checked against published test vectors and against other implementations (Node's crypto; Chrome's PDF viewer).
- **Restrictions are not encryption.** Word document protection, editable ranges, locked regions, PDF permission flags and `AccessPolicy` stop edits in Word, PDF viewers and docxkeun. They do not stop someone who edits the file directly. If your server accepts edited documents, check permissions there.
