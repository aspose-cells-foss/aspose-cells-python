"""
Compound File Binary (CFB) Format Handler

This module handles reading and writing CFB (also known as OLE) format files
used for encrypted XLSX packages according to ECMA-376.
"""

import struct
import base64
import xml.etree.ElementTree as ET
from pathlib import Path

from .encryption_params import HashAlgorithm, CipherAlgorithm, EncryptionType
from .cfb_writer import CFBWriter as CFBWriterImpl


class CFBReader:
    """
    Reads encrypted XLSX from CFB format.
    """

    def __init__(self, file_path):
        """
        Initialize CFB reader.

        Args:
            file_path: Path to CFB file
        """
        self.file_path = file_path
        self._fh = open(file_path, 'rb')
        self._load_header()
        self._load_fat()
        self._load_directory()

    def read_encryption_info(self):
        """
        Read EncryptionInfo stream from CFB file.

        Returns:
            dict with encryption parameters
        """
        data = self._read_stream('EncryptionInfo')
        if data is None:
            raise ValueError("EncryptionInfo stream not found in CFB file")

        # Strip null padding from sector/mini-sector alignment.
        data = data.rstrip(b'\x00')

        # Read version info
        version_major, version_minor = struct.unpack('<HH', data[0:4])
        flags = struct.unpack('<I', data[4:8])[0]

        if version_major == 4 and version_minor == 4:
            # Agile Encryption
            return self._parse_agile_encryption_info(data[8:])
        elif version_major in (2, 3, 4) and version_minor == 2:
            # Standard Encryption
            return self._parse_standard_encryption_info(data, flags)
        else:
            raise ValueError(f"Unsupported encryption version: {version_major}.{version_minor}")

    def _parse_agile_encryption_info(self, xml_data):
        """
        Parse Agile encryption info from XML.

        Args:
            xml_data: XML descriptor bytes

        Returns:
            dict with encryption parameters
        """
        # Parse XML
        root = ET.fromstring(xml_data.decode('utf-8'))

        # Define namespace
        ns = {'enc': 'http://schemas.microsoft.com/office/2006/encryption',
              'p': 'http://schemas.microsoft.com/office/2006/keyEncryptor/password'}

        # Extract keyData
        key_data = root.find('enc:keyData', ns)
        if key_data is None:
            raise ValueError("keyData element not found in Agile encryption info")

        cipher_algorithm_name = key_data.get('cipherAlgorithm', 'AES')
        key_bits = int(key_data.get('keyBits', '256'))
        hash_algorithm_name = key_data.get('hashAlgorithm', 'SHA512')
        salt_value = base64.b64decode(key_data.get('saltValue', ''))

        # Extract encryptedKey from keyEncryptors
        encrypted_key_elem = root.find('.//p:encryptedKey', ns)
        if encrypted_key_elem is None:
            raise ValueError("encryptedKey element not found")

        spin_count = int(encrypted_key_elem.get('spinCount', '100000'))
        key_salt = base64.b64decode(encrypted_key_elem.get('saltValue', ''))
        encrypted_verifier = base64.b64decode(encrypted_key_elem.get('encryptedVerifierHashInput', ''))
        encrypted_verifier_hash = base64.b64decode(encrypted_key_elem.get('encryptedVerifierHashValue', ''))
        encrypted_key_value = base64.b64decode(encrypted_key_elem.get('encryptedKeyValue', ''))

        data_integrity = root.find('enc:dataIntegrity', ns)
        encrypted_hmac_key = b''
        encrypted_hmac_value = b''
        if data_integrity is not None:
            encrypted_hmac_key = base64.b64decode(data_integrity.get('encryptedHmacKey', '') or b'')
            encrypted_hmac_value = base64.b64decode(data_integrity.get('encryptedHmacValue', '') or b'')

        # Map algorithm names to enums
        cipher_algorithm = self._map_cipher_algorithm(cipher_algorithm_name, key_bits)
        hash_algorithm = self._map_hash_algorithm(hash_algorithm_name)

        return {
            'type': EncryptionType.AGILE,
            'cipher_algorithm': cipher_algorithm,
            'hash_algorithm': hash_algorithm,
            'spin_count': spin_count,
            'salt': key_salt,
            'encrypted_verifier': encrypted_verifier,
            'encrypted_verifier_hash': encrypted_verifier_hash,
            'encrypted_key_value': encrypted_key_value,
            'encrypted_hmac_key': encrypted_hmac_key,
            'encrypted_hmac_value': encrypted_hmac_value,
            'package_salt': salt_value
        }

    def _parse_standard_encryption_info(self, data, flags):
        """
        Parse Standard encryption info from binary data.

        Args:
            data: Full EncryptionInfo data
            flags: Flags value

        Returns:
            dict with encryption parameters
        """
        # This is a simplified parser for Standard encryption
        # Full implementation would need to parse the complete binary structure
        raise NotImplementedError("Standard encryption is not yet supported. "
                                "Please use Agile encryption (Office 2010+)")

    def _map_cipher_algorithm(self, name, key_bits):
        """Map cipher algorithm name and key bits to enum."""
        if name == 'AES':
            if key_bits == 128:
                return CipherAlgorithm.AES_128
            elif key_bits == 192:
                return CipherAlgorithm.AES_192
            elif key_bits == 256:
                return CipherAlgorithm.AES_256
        raise ValueError(f"Unsupported cipher algorithm: {name}-{key_bits}")

    def _map_hash_algorithm(self, name):
        """Map hash algorithm name to enum."""
        name_upper = name.upper().replace('-', '')
        if name_upper == 'SHA1':
            return HashAlgorithm.SHA1
        elif name_upper == 'SHA256':
            return HashAlgorithm.SHA256
        elif name_upper == 'SHA384':
            return HashAlgorithm.SHA384
        elif name_upper == 'SHA512':
            return HashAlgorithm.SHA512
        raise ValueError(f"Unsupported hash algorithm: {name}")

    def read_encrypted_package(self):
        """
        Read EncryptedPackage stream from CFB file.

        Returns:
            bytes: Encrypted package data
        """
        data = self._read_stream_raw('EncryptedPackage')
        if data is None:
            raise ValueError("EncryptedPackage stream not found in CFB file")

        # Read package size (first 8 bytes as uint64)
        size_bytes = data[:8]
        package_size = struct.unpack('<Q', size_bytes)[0]
        encrypted_payload = data[8:]
        return package_size, encrypted_payload, data

    def close(self):
        """Close the CFB file."""
        if self._fh:
            self._fh.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def _load_header(self):
        self._fh.seek(0)
        header = self._fh.read(512)
        sig = struct.unpack('<Q', header[0:8])[0]
        if sig != 0xE11AB1A1E011CFD0:
            raise ValueError("Invalid CFB signature")

        self.sector_shift = struct.unpack('<H', header[30:32])[0]
        self.sector_size = 1 << self.sector_shift
        self.mini_sector_size = 1 << struct.unpack('<H', header[32:34])[0]
        self.num_fat_sectors = struct.unpack('<I', header[44:48])[0]
        self.dir_start = struct.unpack('<I', header[48:52])[0]
        self.mini_stream_cutoff = struct.unpack('<I', header[56:60])[0]
        self.mini_fat_start = struct.unpack('<I', header[60:64])[0]
        self.num_mini_fat_sectors = struct.unpack('<I', header[64:68])[0]

        # DIFAT entries (first 109)
        self.difat = []
        for i in range(109):
            entry = struct.unpack('<I', header[76 + i * 4: 80 + i * 4])[0]
            if entry != 0xFFFFFFFF:
                self.difat.append(entry)

    def _read_sector(self, sector_index):
        self._fh.seek(512 + sector_index * self.sector_size)
        return self._fh.read(self.sector_size)

    def _load_fat(self):
        fat_entries = []
        for i in range(self.num_fat_sectors):
            sector = self.difat[i]
            data = self._read_sector(sector)
            for j in range(0, len(data), 4):
                fat_entries.append(struct.unpack('<I', data[j:j + 4])[0])
        self.fat = fat_entries

    def _load_directory(self):
        dir_data = self._read_stream_by_chain(self.dir_start, is_mini=False)
        entries = []
        for i in range(0, len(dir_data), 128):
            entry = dir_data[i:i + 128]
            if len(entry) < 128:
                break
            name_len = struct.unpack('<H', entry[64:66])[0]
            name_bytes = entry[0:name_len - 2] if name_len >= 2 else b''
            name = name_bytes.decode('utf-16le', errors='ignore')
            obj_type = entry[66]
            color = entry[67]
            left = struct.unpack('<I', entry[68:72])[0]
            right = struct.unpack('<I', entry[72:76])[0]
            child = struct.unpack('<I', entry[76:80])[0]
            starting_sector = struct.unpack('<I', entry[116:120])[0]
            size = struct.unpack('<Q', entry[120:128])[0]
            entries.append({
                'name': name,
                'type': obj_type,
                'color': color,
                'left': left,
                'right': right,
                'child': child,
                'start': starting_sector,
                'size': size,
            })
        self.dir_entries = entries

        # Build name->entry index mapping with path traversal
        self.name_map = {}
        def visit(index, parent_path):
            if index == 0xFFFFFFFF or index >= len(self.dir_entries):
                return
            entry = self.dir_entries[index]
            visit(entry['left'], parent_path)
            name = entry['name']
            path = name if not parent_path else parent_path + '/' + name
            self.name_map[path] = index
            if entry['type'] in (1, 5):  # storage or root
                visit(entry['child'], path if entry['type'] == 1 else "")
            visit(entry['right'], parent_path)
        visit(0, "")

    def _read_stream(self, name):
        index = self.name_map.get(name)
        if index is None:
            return None
        entry = self.dir_entries[index]
        size = entry['size']
        if size == 0:
            return b''
        if size < self.mini_stream_cutoff and entry['type'] == 2:
            return self._read_mini_stream(entry['start'], size)
        return self._read_stream_by_chain(entry['start'], is_mini=False, size=size)

    def _read_stream_raw(self, name):
        index = self.name_map.get(name)
        if index is None:
            return None
        entry = self.dir_entries[index]
        size = entry['size']
        if size == 0:
            return b''
        if size < self.mini_stream_cutoff and entry['type'] == 2:
            return self._read_mini_stream(entry['start'], size)
        data = self._read_stream_by_chain(entry['start'], is_mini=False, size=None)
        return data[:size]

    def _read_stream_by_chain(self, start_sector, is_mini=False, size=None):
        data = bytearray()
        sector = start_sector
        while sector not in (0xFFFFFFFE, 0xFFFFFFFF):
            data.extend(self._read_sector(sector))
            sector = self.fat[sector]
        if size is not None:
            return bytes(data[:size])
        return bytes(data)

    def _load_minifat(self):
        if self.num_mini_fat_sectors == 0:
            self.mini_fat = []
            return
        data = bytearray()
        sector = self.mini_fat_start
        for _ in range(self.num_mini_fat_sectors):
            data.extend(self._read_sector(sector))
            sector = self.fat[sector]
        self.mini_fat = []
        for i in range(0, len(data), 4):
            self.mini_fat.append(struct.unpack('<I', data[i:i + 4])[0])

    def _read_mini_stream(self, start_mini_sector, size):
        if not hasattr(self, 'mini_fat'):
            self._load_minifat()
        # root entry is 0
        root = self.dir_entries[0]
        mini_stream = self._read_stream_by_chain(root['start'], is_mini=False, size=root['size'])
        data = bytearray()
        mini_sector = start_mini_sector
        while mini_sector not in (0xFFFFFFFE, 0xFFFFFFFF):
            offset = mini_sector * self.mini_sector_size
            data.extend(mini_stream[offset:offset + self.mini_sector_size])
            mini_sector = self.mini_fat[mini_sector]
        return bytes(data[:size])


class CFBWriter:
    """
    Writes encrypted XLSX to CFB format.
    """

    def __init__(self):
        """Initialize CFB writer."""
        pass

    def write(self, file_path, encryption_info_xml, encrypted_package, package_size):
        """
        Write CFB file with EncryptionInfo and EncryptedPackage streams.

        Args:
            file_path: Output file path
            encryption_info_xml: EncryptionInfo XML bytes
            encrypted_package: Encrypted package bytes
            package_size: Original (unencrypted) package size in bytes
        """
        # Prepare stream data
        # EncryptionInfo: Version + Flags + XML
        version_info = struct.pack('<HH', 4, 4)  # Major=4, Minor=4
        flags = struct.pack('<I', 0x40)
        encryption_info_data = version_info + flags + encryption_info_xml

        # EncryptedPackage: Size + Data
        package_size_bytes = struct.pack('<Q', package_size)
        encrypted_package_data = package_size_bytes + encrypted_package

        # Create CFB file using MS-CFB compliant writer
        # Use 512-byte sectors (version 3) for maximum compatibility with olefile
        writer = CFBWriterImpl(sector_size=512)
        # Allow CFB writer to place small streams in the mini stream.
        writer.add_stream('EncryptionInfo', encryption_info_data)
        writer.add_stream('EncryptedPackage', encrypted_package_data)
        for name, data in _build_dataspaces_streams().items():
            writer.add_stream(name, data)
        writer.write(file_path)


def _build_dataspaces_streams():
    """
    Build DataSpaces streams required by Office for Agile encryption.
    """
    def _unicode_lp_p4(text):
        data = text.encode('utf-16le')
        return struct.pack('<I', len(data)) + data

    def _pad4(data):
        pad = (-len(data)) % 4
        if pad:
            return data + (b'\x00' * pad)
        return data

    # Version stream
    version_name = "Microsoft.Container.DataSpaces"
    version_data = _unicode_lp_p4(version_name) + struct.pack('<III', 1, 1, 1)

    # DataSpaceMap stream
    component_name = "EncryptedPackage"
    dataspace_name = "StrongEncryptionDataSpace"
    entry = struct.pack('<I', 1)  # cRefComponents
    entry += struct.pack('<I', 0)  # ReferenceComponentType = Stream (0)
    entry += _unicode_lp_p4(component_name)
    entry += _unicode_lp_p4(dataspace_name)
    entry = _pad4(entry)
    # Excel sets entry length to 104 even though entry payload is 100 bytes.
    # Match that for compatibility.
    data_space_map = struct.pack('<II', 8, 1) + struct.pack('<I', len(entry) + 4) + entry

    # DataSpaceInfo stream
    transform_name = "StrongEncryptionTransform"
    ds_info = struct.pack('<II', 8, 1) + _pad4(_unicode_lp_p4(transform_name))

    # TransformInfo Primary stream
    transform_id = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
    transform_type = "Microsoft.Container.EncryptionTransform"
    primary = struct.pack('<II', 88, 1)
    primary += _unicode_lp_p4(transform_id)
    primary += _unicode_lp_p4(transform_type)
    primary += bytes.fromhex(
        "000001000000010000000100000000000000000000000000000004000000"
    )

    return {
        "\x06DataSpaces/Version": version_data,
        "\x06DataSpaces/DataSpaceMap": data_space_map,
        "\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace": ds_info,
        "\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary": primary,
    }


def is_encrypted_file(file_path):
    """
    Check if a file is an encrypted XLSX (CFB format).

    Args:
        file_path: Path to file

    Returns:
        bool: True if file is encrypted (CFB format)
    """
    try:
        with open(file_path, 'rb') as f:
            # Check for CFB signature (first 8 bytes)
            header = f.read(8)
            # CFB signature: D0 CF 11 E0 A1 B1 1A E1
            return header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    except Exception:
        return False
