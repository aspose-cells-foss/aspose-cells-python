"""
XLSX File Encryption and Decryption

This module provides high-level encryption and decryption functionality for XLSX files
according to ECMA-376 Part 2 specification.
"""

import io
import os
import base64
import hmac
import struct
import tempfile
import xml.etree.ElementTree as ET
from Crypto.Cipher import AES

from .encryption_params import (
    get_default_encryption_params,
    AgileEncryptionParameters,
    EncryptionType
)
from .encryption_crypto import (
    EncryptionVerifier,
    PackageEncryption,
    PasswordDerivation
)
from .cfb_handler import CFBReader, CFBWriter, is_encrypted_file


class XLSXEncryptor:
    """
    Handles encryption of XLSX files.
    """

    def __init__(self, encryption_params=None):
        """
        Initialize XLSX encryptor.

        Args:
            encryption_params: EncryptionParameters object (default: Agile AES-256/SHA-512)
        """
        self.params = encryption_params or get_default_encryption_params()

    def encrypt_file(self, input_xlsx_path, output_path, password):
        """
        Encrypt an XLSX file.

        Args:
            input_xlsx_path: Path to unencrypted XLSX file
            output_path: Path for encrypted output file
            password: Encryption password
        """
        # Read the XLSX package
        with open(input_xlsx_path, 'rb') as f:
            package_data = f.read()

        # Encrypt the package
        encrypted_package = self._encrypt_package(package_data, password)

        # Create encryption info
        encryption_info_xml = self._create_encryption_info(password)

        # Write CFB file
        writer = CFBWriter()
        writer.write(output_path, encryption_info_xml, encrypted_package, len(package_data))

    def _encrypt_package(self, data, password):
        """
        Encrypt package data.

        Args:
            data: Package data bytes
            password: Encryption password

        Returns:
            Encrypted package bytes
        """
        if self.params.encryption_type != EncryptionType.AGILE:
            raise NotImplementedError("Only Agile encryption is currently supported")

        # Generate salts
        key_salt = os.urandom(self.params.salt_size)
        self.package_salt = os.urandom(self.params.salt_size)

        # Generate verifier and encryption key
        verifier_data = EncryptionVerifier.generate_verifier_agile(
            password,
            key_salt,
            self.params.hash_algorithm,
            self.params.cipher_algorithm,
            self.params.spin_count
        )

        # Store verifier data for encryption info
        self.verifier_data = verifier_data

        # Encrypt package (Agile)
        encrypted = PackageEncryption.encrypt_package_agile(
            data,
            verifier_data['key_value'],
            self.package_salt,
            self.params.hash_algorithm,
            self.params.block_size
        )

        # Data integrity (HMAC) per MS-OFFCRYPTO 2.3.4.14
        hmac_key = os.urandom(self.params.hash_algorithm.hash_bytes)
        encrypted_stream = struct.pack('<Q', len(data)) + encrypted
        hmac_value = hmac.new(
            hmac_key,
            encrypted_stream,
            self.params.hash_algorithm.algorithm_name.lower()
        ).digest()

        iv_key = PasswordDerivation.derive_iv_agile(
            self.package_salt,
            EncryptionVerifier.BLOCK_KEY_DATA_INTEGRITY_KEY,
            self.params.hash_algorithm,
            self.params.block_size
        )
        iv_val = PasswordDerivation.derive_iv_agile(
            self.package_salt,
            EncryptionVerifier.BLOCK_KEY_DATA_INTEGRITY_VALUE,
            self.params.hash_algorithm,
            self.params.block_size
        )

        hmac_key_padded = EncryptionVerifier._pad_zero(hmac_key, self.params.block_size)
        hmac_val_padded = EncryptionVerifier._pad_zero(hmac_value, self.params.block_size)
        self.encrypted_hmac_key = AES.new(verifier_data['key_value'], AES.MODE_CBC, iv_key).encrypt(hmac_key_padded)
        self.encrypted_hmac_value = AES.new(verifier_data['key_value'], AES.MODE_CBC, iv_val).encrypt(hmac_val_padded)

        return encrypted


    def _create_encryption_info(self, password):
        """
        Create EncryptionInfo XML descriptor for Agile encryption.

        Returns:
            XML bytes
        """
        # Register namespaces
        ET.register_namespace('', 'http://schemas.microsoft.com/office/2006/encryption')
        ET.register_namespace('p', 'http://schemas.microsoft.com/office/2006/keyEncryptor/password')
        ET.register_namespace('c', 'http://schemas.microsoft.com/office/2006/keyEncryptor/certificate')

        # Create XML structure with proper namespace
        ns_enc = 'http://schemas.microsoft.com/office/2006/encryption'
        ns_p = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password'

        encryption = ET.Element('{%s}encryption' % ns_enc)

        # keyData element
        key_data = ET.SubElement(encryption, 'keyData')
        key_data.set('saltSize', str(self.params.salt_size))
        key_data.set('blockSize', str(self.params.block_size))
        key_data.set('keyBits', str(self.params.key_bits))
        key_data.set('hashSize', str(self.params.hash_algorithm.hash_bytes))
        key_data.set('cipherAlgorithm', self.params.cipher_algorithm.algorithm_name)
        key_data.set('cipherChaining', 'ChainingModeCBC')
        key_data.set('hashAlgorithm', self.params.hash_algorithm.algorithm_name)
        key_data.set('saltValue', base64.b64encode(self.package_salt).decode('ascii'))

        # dataIntegrity element
        data_integrity = ET.SubElement(encryption, 'dataIntegrity')
        data_integrity.set('encryptedHmacKey', base64.b64encode(self.encrypted_hmac_key).decode('ascii'))
        data_integrity.set('encryptedHmacValue', base64.b64encode(self.encrypted_hmac_value).decode('ascii'))

        # keyEncryptors element
        key_encryptors = ET.SubElement(encryption, 'keyEncryptors')
        key_encryptor = ET.SubElement(key_encryptors, 'keyEncryptor',
                                      uri='http://schemas.microsoft.com/office/2006/keyEncryptor/password')

        # encryptedKey element with namespace
        encrypted_key = ET.SubElement(key_encryptor, '{%s}encryptedKey' % ns_p)
        encrypted_key.set('spinCount', str(self.params.spin_count))
        encrypted_key.set('saltSize', str(self.params.salt_size))
        encrypted_key.set('blockSize', str(self.params.block_size))
        encrypted_key.set('keyBits', str(self.params.key_bits))
        encrypted_key.set('hashSize', str(self.params.hash_algorithm.hash_bytes))
        encrypted_key.set('cipherAlgorithm', self.params.cipher_algorithm.algorithm_name)
        encrypted_key.set('cipherChaining', 'ChainingModeCBC')
        encrypted_key.set('hashAlgorithm', self.params.hash_algorithm.algorithm_name)
        encrypted_key.set('saltValue', base64.b64encode(self.verifier_data['verifier_salt']).decode('ascii'))
        encrypted_key.set('encryptedVerifierHashInput',
                         base64.b64encode(self.verifier_data['encrypted_verifier']).decode('ascii'))
        encrypted_key.set('encryptedVerifierHashValue',
                         base64.b64encode(self.verifier_data['encrypted_verifier_hash']).decode('ascii'))

        encrypted_key.set('encryptedKeyValue',
                          base64.b64encode(self.verifier_data['encrypted_key_value']).decode('ascii'))

        # Build XML matching Excel's formatting for compatibility
        xml_str = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" '
            'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" '
            'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">'
            f'<keyData saltSize="{self.params.salt_size}" blockSize="{self.params.block_size}" '
            f'keyBits="{self.params.key_bits}" hashSize="{self.params.hash_algorithm.hash_bytes}" '
            f'cipherAlgorithm="{self.params.cipher_algorithm.algorithm_name}" '
            'cipherChaining="ChainingModeCBC" '
            f'hashAlgorithm="{self.params.hash_algorithm.algorithm_name}" '
            f'saltValue="{base64.b64encode(self.package_salt).decode("ascii")}"/>'
            f'<dataIntegrity encryptedHmacKey="{base64.b64encode(self.encrypted_hmac_key).decode("ascii")}" '
            f'encryptedHmacValue="{base64.b64encode(self.encrypted_hmac_value).decode("ascii")}"/>'
            '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
            f'<p:encryptedKey spinCount="{self.params.spin_count}" saltSize="{self.params.salt_size}" '
            f'blockSize="{self.params.block_size}" keyBits="{self.params.key_bits}" '
            f'hashSize="{self.params.hash_algorithm.hash_bytes}" '
            f'cipherAlgorithm="{self.params.cipher_algorithm.algorithm_name}" '
            'cipherChaining="ChainingModeCBC" '
            f'hashAlgorithm="{self.params.hash_algorithm.algorithm_name}" '
            f'saltValue="{base64.b64encode(self.verifier_data["verifier_salt"]).decode("ascii")}" '
            f'encryptedVerifierHashInput="{base64.b64encode(self.verifier_data["encrypted_verifier"]).decode("ascii")}" '
            f'encryptedVerifierHashValue="{base64.b64encode(self.verifier_data["encrypted_verifier_hash"]).decode("ascii")}" '
            f'encryptedKeyValue="{base64.b64encode(self.verifier_data["encrypted_key_value"]).decode("ascii")}"/>'
            '</keyEncryptor></keyEncryptors></encryption>'
        )
        return xml_str.encode('utf-8')


class XLSXDecryptor:
    """
    Handles decryption of XLSX files.
    """

    def decrypt_file(self, input_encrypted_path, output_path, password):
        """
        Decrypt an encrypted XLSX file.

        Args:
            input_encrypted_path: Path to encrypted XLSX file (CFB format)
            output_path: Path for decrypted XLSX output
            password: Decryption password

        Returns:
            bool: True if successful

        Raises:
            ValueError: If password is incorrect or file format is invalid
        """
        # Open CFB file
        with CFBReader(input_encrypted_path) as reader:
            # Read encryption info
            enc_info = reader.read_encryption_info()

            # Read encrypted package
            package_size, encrypted_package, raw_stream = reader.read_encrypted_package()

        # Decrypt the package
        decrypted_package = self._decrypt_package(
            encrypted_package,
            package_size,
            password,
            enc_info,
            raw_stream
        )

        if decrypted_package is None:
            raise ValueError("Incorrect password or corrupted file")

        # Write decrypted XLSX
        with open(output_path, 'wb') as f:
            f.write(decrypted_package)

        return True

    def decrypt_to_memory(self, input_encrypted_path, password):
        """
        Decrypt an encrypted XLSX file to memory.

        Args:
            input_encrypted_path: Path to encrypted XLSX file
            password: Decryption password

        Returns:
            BytesIO: Decrypted XLSX data in memory

        Raises:
            ValueError: If password is incorrect or file format is invalid
        """
        # Create temporary file for decrypted data
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp_path = tmp.name

        try:
            # Decrypt to temp file
            self.decrypt_file(input_encrypted_path, tmp_path, password)

            # Read into memory
            with open(tmp_path, 'rb') as f:
                data = f.read()

            return io.BytesIO(data)

        finally:
            # Clean up temp file
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def _decrypt_package(self, encrypted_data, package_size, password, enc_info, encrypted_stream):
        """
        Decrypt package data.

        Args:
            encrypted_data: Encrypted package bytes
            package_size: Original package size
            password: Decryption password
            enc_info: Encryption info dict

        Returns:
            Decrypted package bytes or None if password is incorrect
        """
        if enc_info['type'] != EncryptionType.AGILE:
            raise NotImplementedError("Only Agile encryption is currently supported")

        # Verify password and get password hash
        password_hash = EncryptionVerifier.verify_password_agile(
            password,
            enc_info['salt'],
            enc_info['encrypted_verifier'],
            enc_info['encrypted_verifier_hash'],
            enc_info['hash_algorithm'],
            enc_info['cipher_algorithm'],
            enc_info['spin_count']
        )

        if password_hash is None:
            return None

        # Decrypt intermediate key value
        secret_key = EncryptionVerifier.decrypt_key_value(
            password_hash,
            enc_info['salt'],
            enc_info['encrypted_key_value'],
            enc_info['hash_algorithm'],
            enc_info['cipher_algorithm']
        )

        # Verify data integrity if available
        if enc_info.get('encrypted_hmac_key') and enc_info.get('encrypted_hmac_value'):
            hmac_key, hmac_value = EncryptionVerifier.decrypt_data_integrity(
                secret_key,
                enc_info['package_salt'],
                enc_info['encrypted_hmac_key'],
                enc_info['encrypted_hmac_value'],
                enc_info['hash_algorithm'],
                enc_info['cipher_algorithm'].block_size,
                enc_info['hash_algorithm'].hash_bytes
            )
            calc = hmac.new(
                hmac_key,
                encrypted_stream,
                enc_info['hash_algorithm'].algorithm_name.lower()
            ).digest()
            if hmac_value[:len(calc)] != calc:
                raise ValueError("Data integrity check failed")

        # Decrypt package (Agile)
        decrypted = PackageEncryption.decrypt_package_agile(
            encrypted_data,
            secret_key,
            enc_info['package_salt'],
            enc_info['hash_algorithm'],
            enc_info['cipher_algorithm'].block_size
        )

        return decrypted[:package_size]


def encrypt_xlsx(input_path, output_path, password, encryption_params=None):
    """
    Convenience function to encrypt an XLSX file.

    Args:
        input_path: Path to unencrypted XLSX file
        output_path: Path for encrypted output
        password: Encryption password
        encryption_params: Optional EncryptionParameters (default: AES-256/SHA-512)
    """
    encryptor = XLSXEncryptor(encryption_params)
    encryptor.encrypt_file(input_path, output_path, password)


def decrypt_xlsx(input_path, output_path, password):
    """
    Convenience function to decrypt an XLSX file.

    Args:
        input_path: Path to encrypted XLSX file
        output_path: Path for decrypted output
        password: Decryption password

    Returns:
        bool: True if successful

    Raises:
        ValueError: If password is incorrect
    """
    decryptor = XLSXDecryptor()
    return decryptor.decrypt_file(input_path, output_path, password)
