"""
XLSX Encryption Cryptographic Operations

This module implements cryptographic operations for XLSX encryption/decryption
according to MS-OFFCRYPTO (Agile Encryption).
"""

import hashlib
import os
import struct
from Crypto.Cipher import AES


class PasswordDerivation:
    """Password derivation helpers for Agile encryption."""

    @staticmethod
    def derive_hash_agile(password, salt, hash_algorithm, spin_count):
        """
        Derive H_n for Agile encryption (MS-OFFCRYPTO 2.3.4.11).

        H_0 = Hash(salt + password)
        H_i = Hash(i + H_{i-1}) for i in [0, spin_count)
        """
        password_bytes = password.encode('utf-16le')
        hash_func_name = hash_algorithm.algorithm_name.lower()

        h = hashlib.new(hash_func_name)
        h.update(salt)
        h.update(password_bytes)
        current = h.digest()

        for i in range(spin_count):
            h = hashlib.new(hash_func_name)
            h.update(struct.pack('<I', i))
            h.update(current)
            current = h.digest()

        return current

    @staticmethod
    def derive_key_with_block_key(h_base, block_key, hash_algorithm, key_bits):
        """
        Derive key from H_n and block key (MS-OFFCRYPTO 2.3.4.11).
        """
        hash_func_name = hash_algorithm.algorithm_name.lower()
        h = hashlib.new(hash_func_name)
        h.update(h_base)
        h.update(block_key)
        derived = h.digest()

        key_bytes = key_bits // 8
        if len(derived) >= key_bytes:
            return derived[:key_bytes]
        return derived + (b'\x36' * (key_bytes - len(derived)))

    @staticmethod
    def derive_iv_agile(salt, block_key, hash_algorithm, block_size):
        """
        Derive IV for Agile encryption (MS-OFFCRYPTO 2.3.4.12).

        If block_key is None, IV = salt (padded/truncated to block_size).
        Otherwise IV = Hash(salt + block_key), padded/truncated to block_size.
        """
        if block_key is None:
            iv = salt
        else:
            hash_func_name = hash_algorithm.algorithm_name.lower()
            h = hashlib.new(hash_func_name)
            h.update(salt)
            h.update(block_key)
            iv = h.digest()

        if len(iv) >= block_size:
            return iv[:block_size]
        return iv + (b'\x36' * (block_size - len(iv)))


class EncryptionVerifier:
    """Encryption verifier generation and validation."""

    BLOCK_KEY_VERIFIER = bytes.fromhex('fea7d2763b4b9e79')
    BLOCK_KEY_VERIFIER_HASH = bytes.fromhex('d7aa0f6d3061344e')
    BLOCK_KEY_KEYVALUE = bytes.fromhex('146e0be7abacd0d6')
    BLOCK_KEY_DATA_INTEGRITY_KEY = bytes.fromhex('5fb2ad010cb9e1f6')
    BLOCK_KEY_DATA_INTEGRITY_VALUE = bytes.fromhex('a0677f02b22c8433')

    @staticmethod
    def _pad_zero(data, block_size):
        if len(data) % block_size == 0:
            return data
        pad_len = block_size - (len(data) % block_size)
        return data + (b'\x00' * pad_len)

    @staticmethod
    def generate_verifier_agile(password, salt, hash_algorithm, cipher_algorithm, spin_count):
        """
        Generate PasswordKeyEncryptor fields (MS-OFFCRYPTO 2.3.4.13).
        """
        verifier_input = os.urandom(16)
        hash_func_name = hash_algorithm.algorithm_name.lower()
        verifier_hash = hashlib.new(hash_func_name, verifier_input).digest()

        h_base = PasswordDerivation.derive_hash_agile(
            password, salt, hash_algorithm, spin_count
        )

        verifier_key = PasswordDerivation.derive_key_with_block_key(
            h_base, EncryptionVerifier.BLOCK_KEY_VERIFIER,
            hash_algorithm, cipher_algorithm.key_bits
        )
        iv = PasswordDerivation.derive_iv_agile(
            salt, None, hash_algorithm, cipher_algorithm.block_size
        )
        encrypted_verifier = AES.new(verifier_key, AES.MODE_CBC, iv).encrypt(verifier_input)

        hash_key = PasswordDerivation.derive_key_with_block_key(
            h_base, EncryptionVerifier.BLOCK_KEY_VERIFIER_HASH,
            hash_algorithm, cipher_algorithm.key_bits
        )
        verifier_hash_padded = EncryptionVerifier._pad_zero(
            verifier_hash, cipher_algorithm.block_size
        )
        encrypted_verifier_hash = AES.new(hash_key, AES.MODE_CBC, iv).encrypt(verifier_hash_padded)

        key_value = os.urandom(cipher_algorithm.key_bits // 8)
        key_value_key = PasswordDerivation.derive_key_with_block_key(
            h_base, EncryptionVerifier.BLOCK_KEY_KEYVALUE,
            hash_algorithm, cipher_algorithm.key_bits
        )
        encrypted_key_value = AES.new(key_value_key, AES.MODE_CBC, iv).encrypt(key_value)

        return {
            'verifier_salt': salt,
            'encrypted_verifier': encrypted_verifier,
            'encrypted_verifier_hash': encrypted_verifier_hash,
            'encrypted_key_value': encrypted_key_value,
            'key_value': key_value,
            'password_hash': h_base
        }

    @staticmethod
    def verify_password_agile(password, salt, encrypted_verifier, encrypted_verifier_hash,
                              hash_algorithm, cipher_algorithm, spin_count):
        """
        Verify password for Agile encryption (MS-OFFCRYPTO 2.3.4.13).

        Returns:
            H_n if password is correct, None otherwise.
        """
        try:
            h_base = PasswordDerivation.derive_hash_agile(
                password, salt, hash_algorithm, spin_count
            )

            verifier_key = PasswordDerivation.derive_key_with_block_key(
                h_base, EncryptionVerifier.BLOCK_KEY_VERIFIER,
                hash_algorithm, cipher_algorithm.key_bits
            )
            iv = PasswordDerivation.derive_iv_agile(
                salt, None, hash_algorithm, cipher_algorithm.block_size
            )
            decrypted_verifier = AES.new(verifier_key, AES.MODE_CBC, iv).decrypt(encrypted_verifier)

            hash_func_name = hash_algorithm.algorithm_name.lower()
            computed_hash = hashlib.new(hash_func_name, decrypted_verifier).digest()

            hash_key = PasswordDerivation.derive_key_with_block_key(
                h_base, EncryptionVerifier.BLOCK_KEY_VERIFIER_HASH,
                hash_algorithm, cipher_algorithm.key_bits
            )
            decrypted_hash = AES.new(hash_key, AES.MODE_CBC, iv).decrypt(encrypted_verifier_hash)

            if computed_hash != decrypted_hash[:len(computed_hash)]:
                return None

            return h_base
        except Exception:
            return None

    @staticmethod
    def decrypt_key_value(h_base, key_salt, encrypted_key_value, hash_algorithm, cipher_algorithm):
        """Decrypt intermediate key value (MS-OFFCRYPTO 2.3.4.13)."""
        key_value_key = PasswordDerivation.derive_key_with_block_key(
            h_base, EncryptionVerifier.BLOCK_KEY_KEYVALUE,
            hash_algorithm, cipher_algorithm.key_bits
        )
        iv = PasswordDerivation.derive_iv_agile(
            key_salt, None, hash_algorithm, cipher_algorithm.block_size
        )
        return AES.new(key_value_key, AES.MODE_CBC, iv).decrypt(encrypted_key_value)

    @staticmethod
    def decrypt_data_integrity(secret_key, package_salt, encrypted_hmac_key, encrypted_hmac_value,
                               hash_algorithm, block_size, hash_size):
        """Decrypt HMAC key and value (MS-OFFCRYPTO 2.3.4.14)."""
        iv_key = PasswordDerivation.derive_iv_agile(
            package_salt, EncryptionVerifier.BLOCK_KEY_DATA_INTEGRITY_KEY,
            hash_algorithm, block_size
        )
        iv_val = PasswordDerivation.derive_iv_agile(
            package_salt, EncryptionVerifier.BLOCK_KEY_DATA_INTEGRITY_VALUE,
            hash_algorithm, block_size
        )
        hmac_key = AES.new(secret_key, AES.MODE_CBC, iv_key).decrypt(encrypted_hmac_key)
        hmac_value = AES.new(secret_key, AES.MODE_CBC, iv_val).decrypt(encrypted_hmac_value)
        return hmac_key[:hash_size], hmac_value[:hash_size]


class PackageEncryption:
    """Package data encryption and decryption."""

    @staticmethod
    def encrypt_package_agile(data, key, package_salt, hash_algorithm, block_size):
        """Encrypt package data using Agile encryption (MS-OFFCRYPTO 2.3.4.15)."""
        encrypted = bytearray()
        segment_size = 4096

        for segment_index in range(0, len(data), segment_size):
            chunk = data[segment_index:segment_index + segment_size]
            block_key = struct.pack('<I', segment_index // segment_size)
            iv = PasswordDerivation.derive_iv_agile(
                package_salt, block_key, hash_algorithm, block_size
            )
            if len(chunk) % block_size != 0:
                chunk = chunk + (b'\x00' * (block_size - (len(chunk) % block_size)))
            cipher = AES.new(key, AES.MODE_CBC, iv)
            encrypted.extend(cipher.encrypt(chunk))

        return bytes(encrypted)

    @staticmethod
    def decrypt_package_agile(encrypted_data, key, package_salt, hash_algorithm, block_size):
        """Decrypt package data using Agile encryption (MS-OFFCRYPTO 2.3.4.15)."""
        decrypted = bytearray()
        segment_size = 4096

        for segment_index in range(0, len(encrypted_data), segment_size):
            chunk = encrypted_data[segment_index:segment_index + segment_size]
            block_key = struct.pack('<I', segment_index // segment_size)
            iv = PasswordDerivation.derive_iv_agile(
                package_salt, block_key, hash_algorithm, block_size
            )
            cipher = AES.new(key, AES.MODE_CBC, iv)
            decrypted.extend(cipher.decrypt(chunk))

        return bytes(decrypted)
