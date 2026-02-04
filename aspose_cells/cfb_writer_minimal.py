"""
Minimal CFB (Compound File Binary) Writer

This module provides a minimal CFB file writer for creating encrypted XLSX files.
CFB format is also known as OLE (Object Linking and Embedding) format.

Based on MS-CFB specification.
"""

import struct
import io


class MinimalCFBWriter:
    """
    Minimal CFB file writer for encrypted Office documents.

    This implementation creates a simple CFB file with just two streams:
    - EncryptionInfo
    - EncryptedPackage
    """

    # CFB Constants
    HEADER_SIGNATURE = 0xE11AB1A1E011CFD0
    MINOR_VERSION = 0x003E
    MAJOR_VERSION_3 = 0x0003  # Version 3 (512-byte sectors)
    MAJOR_VERSION_4 = 0x0004  # Version 4 (4096-byte sectors)
    BYTE_ORDER = 0xFFFE
    SECTOR_SIZE_V3 = 512
    SECTOR_SIZE_V4 = 4096
    MINI_SECTOR_SIZE = 64
    DIFSECT = 0xFFFFFFFC
    FATSECT = 0xFFFFFFFD
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF

    def __init__(self, sector_size=512):
        """
        Initialize CFB writer.

        Args:
            sector_size: Sector size (512 or 4096 bytes)
        """
        self.sector_size = sector_size
        self.major_version = self.MAJOR_VERSION_3 if sector_size == 512 else self.MAJOR_VERSION_4
        self.streams = []

    def add_stream(self, name, data):
        """
        Add a stream to the CFB file.

        Args:
            name: Stream name
            data: Stream data bytes
        """
        self.streams.append((name, data))

    def write(self, file_path):
        """
        Write CFB file.

        Args:
            file_path: Output file path
        """
        # Build CFB structure
        header = self._build_header()
        fat = self._build_fat()
        directory = self._build_directory()
        data_sectors = self._build_data_sectors()

        # Write to file
        with open(file_path, 'wb') as f:
            f.write(header)
            f.write(fat)
            f.write(directory)
            f.write(data_sectors)

    def _build_header(self):
        """Build CFB header (512 bytes)."""
        header = io.BytesIO()

        # Signature (8 bytes)
        header.write(struct.pack('<Q', self.HEADER_SIGNATURE))

        # CLSID (16 bytes) - all zeros
        header.write(b'\x00' * 16)

        # Minor version (2 bytes)
        header.write(struct.pack('<H', self.MINOR_VERSION))

        # Major version (2 bytes)
        header.write(struct.pack('<H', self.major_version))

        # Byte order (2 bytes)
        header.write(struct.pack('<H', self.BYTE_ORDER))

        # Sector size power (2 bytes): 9 for 512, 12 for 4096
        sector_shift = 9 if self.sector_size == 512 else 12
        header.write(struct.pack('<H', sector_shift))

        # Mini sector size power (2 bytes): always 6 (64 bytes)
        header.write(struct.pack('<H', 6))

        # Reserved (6 bytes)
        header.write(b'\x00' * 6)

        # Total sectors (4 bytes) - 0 for version 3
        header.write(struct.pack('<I', 0))

        # FAT sectors (4 bytes)
        fat_sectors = self._calculate_fat_sectors()
        header.write(struct.pack('<I', fat_sectors))

        # First directory sector (4 bytes)
        header.write(struct.pack('<I', fat_sectors))  # Right after FAT

        # Transaction signature (4 bytes)
        header.write(struct.pack('<I', 0))

        # Mini stream cutoff size (4 bytes) - 4096
        header.write(struct.pack('<I', 4096))

        # First mini FAT sector (4 bytes) - ENDOFCHAIN (no mini stream)
        header.write(struct.pack('<I', self.ENDOFCHAIN))

        # Number of mini FAT sectors (4 bytes)
        header.write(struct.pack('<I', 0))

        # First DIFAT sector (4 bytes) - ENDOFCHAIN
        header.write(struct.pack('<I', self.ENDOFCHAIN))

        # Number of DIFAT sectors (4 bytes)
        header.write(struct.pack('<I', 0))

        # DIFAT array (436 bytes for 109 entries)
        # First FAT sector is at sector 0
        for i in range(109):
            if i == 0:
                header.write(struct.pack('<I', 0))  # FAT at sector 0
            else:
                header.write(struct.pack('<I', self.FREESECT))

        # Pad to 512 bytes
        data = header.getvalue()
        if len(data) < 512:
            data += b'\x00' * (512 - len(data))

        return data[:512]

    def _calculate_fat_sectors(self):
        """Calculate number of FAT sectors needed."""
        # We need:
        # - 1 FAT sector
        # - 1 directory sector
        # - N data sectors for streams

        total_data = sum(len(data) for _, data in self.streams)
        data_sectors = (total_data + self.sector_size - 1) // self.sector_size

        # For simplicity, allocate 1 FAT sector (can hold 128 entries for 512-byte sectors)
        return 1

    def _build_fat(self):
        """Build File Allocation Table."""
        entries_per_sector = self.sector_size // 4
        fat = io.BytesIO()

        # FAT entries:
        # Sector 0: FAT sector itself (FATSECT)
        fat.write(struct.pack('<I', self.FATSECT))

        # Sector 1: Directory sector (ENDOFCHAIN)
        fat.write(struct.pack('<I', self.ENDOFCHAIN))

        # Build separate chains for each stream
        current_sector = 2
        for stream_idx, (name, data) in enumerate(self.streams):
            sectors_needed = (len(data) + self.sector_size - 1) // self.sector_size

            # Create chain for this stream
            for i in range(sectors_needed):
                if i == sectors_needed - 1:
                    # Last sector of this stream
                    fat.write(struct.pack('<I', self.ENDOFCHAIN))
                else:
                    # Point to next sector in chain
                    fat.write(struct.pack('<I', current_sector + i + 1))

            current_sector += sectors_needed

        # Fill rest of FAT sector with FREESECT
        total_data_sectors = sum((len(d) + self.sector_size - 1) // self.sector_size
                                 for _, d in self.streams)
        entries_written = 2 + total_data_sectors
        for i in range(entries_written, entries_per_sector):
            fat.write(struct.pack('<I', self.FREESECT))

        return fat.getvalue()

    def _build_directory(self):
        """Build directory entries."""
        directory = io.BytesIO()

        # Root entry (entry 0)
        root_entry = self._create_directory_entry(
            name="Root Entry",
            obj_type=5,  # Root storage
            color=1,  # Black
            left_sibling=0xFFFFFFFF,
            right_sibling=0xFFFFFFFF,
            child_did=1 if len(self.streams) > 0 else 0xFFFFFFFF,
            clsid=b'\x00' * 16,
            state_bits=0,
            creation_time=0,
            modified_time=0,
            starting_sector=0xFFFFFFFE,  # ENDOFCHAIN
            stream_size=0
        )
        directory.write(root_entry)

        # Stream entries
        # For simplicity with 2 streams, create simple red-black tree:
        # Stream 0 (EncryptionInfo) - entry 1, left child of root
        # Stream 1 (EncryptedPackage) - entry 2, right sibling of stream 0
        for i, (name, data) in enumerate(self.streams):
            # Calculate starting sector (after FAT and directory)
            if i == 0:
                starting_sector = 2  # After FAT (0) and directory (1)
            else:
                prev_data = sum(len(d) for _, d in self.streams[:i])
                prev_sectors = (prev_data + self.sector_size - 1) // self.sector_size
                starting_sector = 2 + prev_sectors

            # Set up tree structure for exactly 2 streams
            if i == 0:
                # First stream: left child of root, right sibling is second stream
                left_sib = 0xFFFFFFFF
                right_sib = 2 if len(self.streams) > 1 else 0xFFFFFFFF
            elif i == 1:
                # Second stream: right sibling of first stream
                left_sib = 0xFFFFFFFF
                right_sib = 0xFFFFFFFF
            else:
                left_sib = 0xFFFFFFFF
                right_sib = 0xFFFFFFFF

            entry = self._create_directory_entry(
                name=name,
                obj_type=2,  # Stream
                color=1,  # Black
                left_sibling=left_sib,
                right_sibling=right_sib,
                child_did=0xFFFFFFFF,
                clsid=b'\x00' * 16,
                state_bits=0,
                creation_time=0,
                modified_time=0,
                starting_sector=starting_sector,
                stream_size=len(data)
            )
            directory.write(entry)

        # Fill rest of directory sector
        entries_per_sector = self.sector_size // 128
        entries_written = 1 + len(self.streams)
        for i in range(entries_written, entries_per_sector):
            directory.write(b'\xFF' * 128)

        return directory.getvalue()

    def _create_directory_entry(self, name, obj_type, color, left_sibling, right_sibling,
                                child_did, clsid, state_bits, creation_time, modified_time,
                                starting_sector, stream_size):
        """Create a directory entry (128 bytes)."""
        entry = io.BytesIO()

        # Name (64 bytes): UTF-16LE, null-terminated
        name_bytes = name.encode('utf-16le')
        if len(name_bytes) > 62:
            name_bytes = name_bytes[:62]
        entry.write(name_bytes)
        entry.write(b'\x00' * (64 - len(name_bytes)))

        # Name length (2 bytes): including null terminator
        entry.write(struct.pack('<H', len(name_bytes) + 2))

        # Object type (1 byte)
        entry.write(struct.pack('<B', obj_type))

        # Color (1 byte)
        entry.write(struct.pack('<B', color))

        # Left sibling DID (4 bytes)
        entry.write(struct.pack('<I', left_sibling))

        # Right sibling DID (4 bytes)
        entry.write(struct.pack('<I', right_sibling))

        # Child DID (4 bytes)
        entry.write(struct.pack('<I', child_did))

        # CLSID (16 bytes)
        entry.write(clsid)

        # State bits (4 bytes)
        entry.write(struct.pack('<I', state_bits))

        # Creation time (8 bytes)
        entry.write(struct.pack('<Q', creation_time))

        # Modified time (8 bytes)
        entry.write(struct.pack('<Q', modified_time))

        # Starting sector (4 bytes)
        entry.write(struct.pack('<I', starting_sector))

        # Stream size (8 bytes)
        entry.write(struct.pack('<Q', stream_size))

        data = entry.getvalue()
        assert len(data) == 128
        return data

    def _build_data_sectors(self):
        """Build data sectors for all streams."""
        data = io.BytesIO()

        for name, stream_data in self.streams:
            # Write stream data
            data.write(stream_data)

            # Pad to sector boundary
            padding = (self.sector_size - (len(stream_data) % self.sector_size)) % self.sector_size
            if padding > 0:
                data.write(b'\x00' * padding)

        return data.getvalue()
