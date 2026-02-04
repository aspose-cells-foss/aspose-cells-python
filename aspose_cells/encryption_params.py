"""
XLSX Encryption Parameters

This module defines encryption parameter classes for Standard and Agile encryption
according to ECMA-376 Part 2 specification.
"""

from enum import Enum


class EncryptionType(Enum):
    """Encryption type enumeration."""
    STANDARD = "Standard"
    AGILE = "Agile"


class CipherAlgorithm(Enum):
    """Cipher algorithm enumeration."""
    AES_128 = ("AES", 128, 0x660E)
    AES_192 = ("AES", 192, 0x660F)
    AES_256 = ("AES", 256, 0x6610)

    def __init__(self, name, key_bits, alg_id):
        self.algorithm_name = name
        self.key_bits = key_bits
        self.alg_id = alg_id

    @property
    def key_bytes(self):
        """Get key size in bytes."""
        return self.key_bits // 8

    @property
    def block_size(self):
        """Get block size in bytes (AES uses 16-byte blocks)."""
        return 16


class HashAlgorithm(Enum):
    """Hash algorithm enumeration."""
    SHA1 = ("SHA1", 20, 0x8004)
    SHA256 = ("SHA256", 32, 0x800C)
    SHA384 = ("SHA384", 48, 0x800D)
    SHA512 = ("SHA512", 64, 0x800E)

    def __init__(self, name, hash_bytes, alg_id):
        self.algorithm_name = name
        self.hash_bytes = hash_bytes
        self.alg_id = alg_id


class EncryptionParameters:
    """
    Base class for encryption parameters.
    """

    def __init__(self, encryption_type, cipher_algorithm, hash_algorithm, spin_count=100000):
        """
        Initialize encryption parameters.

        Args:
            encryption_type: EncryptionType enum value
            cipher_algorithm: CipherAlgorithm enum value
            hash_algorithm: HashAlgorithm enum value
            spin_count: Number of iterations for key derivation (default: 100,000)
        """
        self.encryption_type = encryption_type
        self.cipher_algorithm = cipher_algorithm
        self.hash_algorithm = hash_algorithm
        self.spin_count = spin_count


class AgileEncryptionParameters(EncryptionParameters):
    """
    Parameters for Agile Encryption (ECMA-376 Part 2, Section 4).

    This is the modern encryption method used by Office 2010+.
    Recommended settings: AES-256 + SHA-512 with 100,000 iterations.
    """

    def __init__(self,
                 cipher_algorithm=CipherAlgorithm.AES_256,
                 hash_algorithm=HashAlgorithm.SHA512,
                 spin_count=100000):
        """
        Initialize Agile encryption parameters.

        Args:
            cipher_algorithm: Cipher algorithm (default: AES-256)
            hash_algorithm: Hash algorithm (default: SHA-512)
            spin_count: Iteration count for key derivation (default: 100,000)
        """
        super().__init__(EncryptionType.AGILE, cipher_algorithm, hash_algorithm, spin_count)

        # Agile encryption specific parameters
        self.salt_size = 16  # 16 bytes for primary salt
        self.block_size = cipher_algorithm.block_size
        self.key_bits = cipher_algorithm.key_bits


class StandardEncryptionParameters(EncryptionParameters):
    """
    Parameters for Standard Encryption (ECMA-376 Part 2, Section 3).

    This is the legacy encryption method used by Office 2007-2009.
    Maintained for compatibility but Agile encryption is recommended.
    """

    def __init__(self,
                 cipher_algorithm=CipherAlgorithm.AES_128,
                 hash_algorithm=HashAlgorithm.SHA1,
                 spin_count=50000):
        """
        Initialize Standard encryption parameters.

        Args:
            cipher_algorithm: Cipher algorithm (default: AES-128)
            hash_algorithm: Hash algorithm (default: SHA-1)
            spin_count: Iteration count for key derivation (default: 50,000)
        """
        super().__init__(EncryptionType.STANDARD, cipher_algorithm, hash_algorithm, spin_count)

        # Standard encryption specific parameters
        self.salt_size = 16  # 16 bytes
        self.verifier_size = 16  # 16 bytes


def get_default_encryption_params():
    """
    Get default encryption parameters.

    Returns Office 2013+ default: Agile encryption with AES-256 and SHA-512.
    """
    return AgileEncryptionParameters(
        cipher_algorithm=CipherAlgorithm.AES_256,
        hash_algorithm=HashAlgorithm.SHA512,
        spin_count=100000
    )
