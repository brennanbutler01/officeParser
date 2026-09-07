# Encrypted-document fixtures

All fixtures here use the password **`test123`**. They come in two kinds.

## Generated (synthetic)

Written by [`scripts/generate-crypto-fixtures.mjs`](../../../scripts/generate-crypto-fixtures.mjs),
**independently of `src/crypto`** (its own from-scratch encryptor), so the tests cross-check the
decryptor against a second implementation rather than against itself. Marker: **`SECRET CONTENT 42`**.

| File | Scheme |
|------|--------|
| `agile.docx` | ECMA-376 **agile** encryption (AES-256-CBC, SHA-512, CFB/OLE2 container) |
| `standard.docx` | ECMA-376 **standard** encryption (AES-ECB, SHA-1, CFB/OLE2 container) |
| `encrypted.odt` | ODF per-entry AES-256-CBC with PBKDF2 |

Regenerate with `node scripts/generate-crypto-fixtures.mjs`.

## Real (LibreOffice output)

Produced by a real **LibreOffice 7.3.7** headless save (its default ODF encryption:
AES-256-CBC, PBKDF2 100000 rounds, SHA-256 start key, W3C random padding), **not** by our own
encryptor. These guard against the decryptor only ever being tested against its sibling encoder:
real writers differ in re-zip order, padding, entry set (directory and 0-byte entries) and content
size. Marker: **`REAL LIBRE MARKER 77`**. Do **not** regenerate these from the script; they are
committed as-is. To refresh, re-save a document containing the marker with LibreOffice 7.x using the
classic AES-CBC/PBKDF2 scheme and password `test123`.

| File | Type |
|------|------|
| `real-libreoffice.odt` | Writer text document |
| `real-libreoffice.ods` | Calc spreadsheet |
| `real-libreoffice.odp` | Impress presentation |

> LibreOffice 24.8+ defaults to AES-256-GCM + Argon2id ("wholesome encryption"), which is
> deliberately unsupported (it rejects with `DOCUMENT_DECRYPTION_FAILED`); these fixtures use the
> classic scheme that officeParser decrypts.
