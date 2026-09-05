/**
 * Document decryption for password-protected inputs, kept out of the format parsers so the parsers
 * only ever see plaintext. The dispatcher detects an encrypted container up front and, when one is
 * found, decrypts it (retrying via `onPassword`) before routing the plaintext to the normal parser.
 *
 * Formats: OOXML (CFB-wrapped, {@link module:crypto/ooxmlCrypto}) and ODF (encrypted zip entries,
 * {@link module:crypto/odfCrypto}). PDF decryption is handled inside pdf.js by the PDF parser, but
 * shares the same top-level `password`/`onPassword` config and the same `PASSWORD_*` error codes.
 *
 * @module crypto
 */

export { isCfb } from './cfb.js';
export { isEncryptedOoxml, decryptOoxml } from './ooxmlCrypto.js';
export { isEncryptedOdf, decryptOdf } from './odfCrypto.js';
export { WRONG_PASSWORD } from './wrongPassword.js';
