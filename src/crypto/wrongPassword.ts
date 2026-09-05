/**
 * Sentinel thrown by a format decryptor when the supplied password does not verify, distinct from a
 * structural/format error. The retry driver in {@link module:crypto/index} catches exactly this to
 * decide between asking `onPassword` again and giving up with `PASSWORD_INCORRECT`.
 *
 * @module crypto/wrongPassword
 */

/** Unique marker error meaning "the password was wrong" (as opposed to "the file is malformed"). */
export const WRONG_PASSWORD: unique symbol = Symbol('officeparser.WRONG_PASSWORD');

/**
 * A structural or unsupported-scheme failure while decrypting (a malformed container, an unsupported
 * cipher/version), distinct from {@link WRONG_PASSWORD}. The dispatcher wraps it into a
 * `DOCUMENT_DECRYPTION_FAILED` OfficeError. A dedicated class (rather than `new Error`) also keeps
 * the low-level crypto modules free of the parser's error/message machinery.
 */
export class DecryptionError extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'DecryptionError';
    }
}
