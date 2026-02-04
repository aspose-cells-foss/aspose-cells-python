"""
CFB (Compound File Binary) Writer

Minimal but compliant implementation for encrypted Office documents.
Supports storages and mini streams (MS-CFB).
"""

import io
import struct


class _Node:
    def __init__(self, name, obj_type, data=None):
        self.name = name
        self.obj_type = obj_type
        self.data = data
        self.children = []
        self.left = None
        self.right = None
        self.child = None
        self.parent = None
        self.color = 1
        self.starting_sector = 0xFFFFFFFE
        self.stream_size = 0 if data is None else len(data)
        self.force_regular = False
        self.manual_tree = False
        self.did = None


class CFBWriter:
    """
    Writes CFB (Compound File Binary) files according to MS-CFB specification.
    """

    HEADER_SIGNATURE = 0xE11AB1A1E011CFD0
    HEADER_CLSID = b'\x00' * 16
    MINOR_VERSION = 0x003E
    MAJOR_VERSION_3 = 0x0003
    MAJOR_VERSION_4 = 0x0004
    BYTE_ORDER = 0xFFFE

    MAXREGSECT = 0xFFFFFFFA
    DIFSECT = 0xFFFFFFFC
    FATSECT = 0xFFFFFFFD
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF

    STGTY_INVALID = 0
    STGTY_STORAGE = 1
    STGTY_STREAM = 2
    STGTY_LOCKBYTES = 3
    STGTY_PROPERTY = 4
    STGTY_ROOT = 5

    COLOR_RED = 0
    COLOR_BLACK = 1

    MINI_STREAM_CUTOFF = 4096
    MINI_SECTOR_SIZE = 64

    def __init__(self, sector_size=512):
        if sector_size not in (512, 4096):
            raise ValueError("Sector size must be 512 or 4096")

        self.sector_size = sector_size
        self.major_version = self.MAJOR_VERSION_3 if sector_size == 512 else self.MAJOR_VERSION_4
        self.sector_shift = 9 if sector_size == 512 else 12
        self.root = _Node("Root Entry", self.STGTY_ROOT)

    def add_stream(self, name, data, force_regular=False):
        if len(name) == 0:
            raise ValueError("Stream name cannot be empty")

        parts = name.split('/')
        node = self.root
        for part in parts[:-1]:
            node = self._get_or_create_storage(node, part)

        stream_node = _Node(parts[-1], self.STGTY_STREAM, data)
        stream_node.force_regular = force_regular
        node.children.append(stream_node)

    def _get_or_create_storage(self, parent, name):
        for child in parent.children:
            if child.obj_type in (self.STGTY_STORAGE, self.STGTY_ROOT) and child.name == name:
                return child
        storage = _Node(name, self.STGTY_STORAGE)
        parent.children.append(storage)
        return storage

    def write(self, file_path):
        self._build_storage_trees(self.root)
        all_nodes = self._collect_nodes()
        self._assign_directory_ids(all_nodes)

        layout = self._calculate_layout(all_nodes)

        header = self._build_header(layout)
        fat_data = self._build_fat(layout)
        fat_sectors = self._split_into_sectors(fat_data)
        dir_data = self._build_directory(all_nodes)
        dir_sectors = self._split_into_sectors(dir_data)
        mini_fat_data = self._build_minifat(layout)
        mini_fat_sectors = self._split_into_sectors(mini_fat_data) if mini_fat_data else []
        stream_sector_map = self._build_stream_sectors(layout)

        sectors = [b'\x00' * self.sector_size for _ in range(layout['total_sectors'])]

        for i, sector_index in enumerate(layout['fat_sectors_list']):
            sectors[sector_index] = fat_sectors[i]

        for i, sector in enumerate(dir_sectors):
            sectors[layout['dir_start'] + i] = sector

        for i, sector in enumerate(mini_fat_sectors):
            sectors[layout['mini_fat_start'] + i] = sector

        for sector_index, data in stream_sector_map.items():
            sectors[sector_index] = data

        with open(file_path, 'wb') as f:
            f.write(header)
            for sector in sectors:
                f.write(sector)

    def _build_storage_trees(self, node):
        if node.obj_type in (self.STGTY_STORAGE, self.STGTY_ROOT):
            if node.children and not node.manual_tree:
                if node.obj_type == self.STGTY_ROOT and self._try_build_excel_tree(node):
                    pass
                else:
                    node.child = self._build_rb_tree(node.children)
            for child in node.children:
                self._build_storage_trees(child)

    def _try_build_excel_tree(self, root):
        names = {child.name: child for child in root.children}
        if "EncryptedPackage" not in names or "EncryptionInfo" not in names or "\x06DataSpaces" not in names:
            return False

        enc_info = names["EncryptionInfo"]
        enc_pkg = names["EncryptedPackage"]
        data_spaces = names["\x06DataSpaces"]

        # Root tree
        root.child = enc_info
        enc_info.left = data_spaces
        enc_info.right = enc_pkg
        enc_pkg.left = None
        enc_pkg.right = None
        data_spaces.left = None
        data_spaces.right = None
        data_spaces.manual_tree = True

        # Colors to match Excel's typical layout
        root.color = self.COLOR_RED
        enc_pkg.color = self.COLOR_RED
        data_spaces.color = self.COLOR_RED
        enc_info.color = self.COLOR_BLACK

        # Build DataSpaces subtree
        ds_children = {child.name: child for child in data_spaces.children}
        required = {"Version", "DataSpaceMap", "DataSpaceInfo", "TransformInfo"}
        if not required.issubset(ds_children.keys()):
            return False

        version = ds_children["Version"]
        data_space_map = ds_children["DataSpaceMap"]
        data_space_info = ds_children["DataSpaceInfo"]
        transform_info = ds_children["TransformInfo"]

        data_spaces.child = data_space_map
        data_space_map.left = version
        data_space_map.right = data_space_info
        version.left = None
        version.right = None
        data_space_info.left = None
        data_space_info.right = transform_info
        transform_info.left = None
        transform_info.right = None

        data_space_map.color = self.COLOR_BLACK
        version.color = self.COLOR_BLACK
        data_space_info.color = self.COLOR_BLACK
        transform_info.color = self.COLOR_RED
        data_space_info.manual_tree = True
        transform_info.manual_tree = True

        # DataSpaceInfo child
        dsi_children = {child.name: child for child in data_space_info.children}
        if "StrongEncryptionDataSpace" in dsi_children:
            strong_ds = dsi_children["StrongEncryptionDataSpace"]
            data_space_info.child = strong_ds
            strong_ds.left = None
            strong_ds.right = None
            strong_ds.color = self.COLOR_BLACK
            strong_ds.manual_tree = True

        # TransformInfo child
        ti_children = {child.name: child for child in transform_info.children}
        if "StrongEncryptionTransform" in ti_children:
            strong_tr = ti_children["StrongEncryptionTransform"]
            transform_info.child = strong_tr
            strong_tr.left = None
            strong_tr.right = None
            strong_tr.color = self.COLOR_BLACK
            strong_tr.manual_tree = True

            st_children = {child.name: child for child in strong_tr.children}
            if "\x06Primary" in st_children:
                primary = st_children["\x06Primary"]
                strong_tr.child = primary
                primary.left = None
                primary.right = None
                primary.color = self.COLOR_BLACK
                primary.manual_tree = True

        return True

    def _build_rb_tree(self, nodes):
        nodes_sorted = sorted(nodes, key=lambda n: (n.name.upper(), n.name))
        root = None
        for node in nodes_sorted:
            node.left = None
            node.right = None
            node.parent = None
            node.color = self.COLOR_RED
            root = self._rb_insert(root, node)
        if root is not None:
            root.color = self.COLOR_BLACK
        return root

    def _compare_names(self, a, b):
        a_up = a.upper()
        b_up = b.upper()
        if a_up < b_up:
            return -1
        if a_up > b_up:
            return 1
        if a < b:
            return -1
        if a > b:
            return 1
        return 0

    def _rb_insert(self, root, node):
        parent = None
        current = root
        while current is not None:
            parent = current
            if self._compare_names(node.name, current.name) < 0:
                current = current.left
            else:
                current = current.right
        node.parent = parent
        if parent is None:
            root = node
        elif self._compare_names(node.name, parent.name) < 0:
            parent.left = node
        else:
            parent.right = node

        return self._rb_insert_fixup(root, node)

    def _rb_insert_fixup(self, root, node):
        while node.parent is not None and node.parent.color == self.COLOR_RED:
            if node.parent.parent is None:
                node.parent.color = self.COLOR_BLACK
                break
            if node.parent == node.parent.parent.left:
                uncle = node.parent.parent.right
                if uncle is not None and uncle.color == self.COLOR_RED:
                    node.parent.color = self.COLOR_BLACK
                    uncle.color = self.COLOR_BLACK
                    node.parent.parent.color = self.COLOR_RED
                    node = node.parent.parent
                else:
                    if node == node.parent.right:
                        node = node.parent
                        root = self._rotate_left(root, node)
                    node.parent.color = self.COLOR_BLACK
                    node.parent.parent.color = self.COLOR_RED
                    root = self._rotate_right(root, node.parent.parent)
            else:
                uncle = node.parent.parent.left
                if uncle is not None and uncle.color == self.COLOR_RED:
                    node.parent.color = self.COLOR_BLACK
                    uncle.color = self.COLOR_BLACK
                    node.parent.parent.color = self.COLOR_RED
                    node = node.parent.parent
                else:
                    if node == node.parent.left:
                        node = node.parent
                        root = self._rotate_right(root, node)
                    node.parent.color = self.COLOR_BLACK
                    node.parent.parent.color = self.COLOR_RED
                    root = self._rotate_left(root, node.parent.parent)
        return root

    def _rotate_left(self, root, x):
        y = x.right
        x.right = y.left
        if y.left is not None:
            y.left.parent = x
        y.parent = x.parent
        if x.parent is None:
            root = y
        elif x == x.parent.left:
            x.parent.left = y
        else:
            x.parent.right = y
        y.left = x
        x.parent = y
        return root

    def _rotate_right(self, root, y):
        x = y.left
        y.left = x.right
        if x.right is not None:
            x.right.parent = y
        x.parent = y.parent
        if y.parent is None:
            root = x
        elif y == y.parent.right:
            y.parent.right = x
        else:
            y.parent.left = x
        x.right = y
        y.parent = x
        return root

    def _collect_nodes(self):
        nodes = []

        def walk(node):
            nodes.append(node)
            for child in node.children:
                walk(child)

        walk(self.root)
        return nodes

    def _assign_directory_ids(self, nodes):
        for idx, node in enumerate(nodes):
            node.did = idx

    def _calculate_layout(self, nodes):
        layout = {}
        entries_per_sector = self.sector_size // 128
        dir_sectors_needed = (len(nodes) + entries_per_sector - 1) // entries_per_sector

        streams = [n for n in nodes if n.obj_type == self.STGTY_STREAM]
        mini_streams = [s for s in streams if s.stream_size < self.MINI_STREAM_CUTOFF and not s.force_regular]
        regular_streams = [s for s in streams if s.stream_size >= self.MINI_STREAM_CUTOFF or s.force_regular]

        # Build mini stream data
        mini_stream_data = b''
        mini_sector_index = 0
        for s in mini_streams:
            s.starting_sector = mini_sector_index
            data = s.data or b''
            s.stream_size = len(data)
            pad = (-len(data)) % self.MINI_SECTOR_SIZE
            mini_stream_data += data + (b'\x00' * pad)
            mini_sector_index += (len(data) + pad) // self.MINI_SECTOR_SIZE

        mini_stream_size = len(mini_stream_data)
        if mini_stream_size and mini_stream_size < 1920:
            pad = 1920 - mini_stream_size
            mini_stream_data += b'\x00' * pad
            mini_stream_size = 1920
            mini_sector_index = mini_stream_size // self.MINI_SECTOR_SIZE
        self.root.stream_size = mini_stream_size

        mini_fat_entries = mini_sector_index
        entries_per_fat_sector = self.sector_size // 4
        mini_fat_sectors = (mini_fat_entries + entries_per_fat_sector - 1) // entries_per_fat_sector
        mini_stream_sectors = (mini_stream_size + self.sector_size - 1) // self.sector_size if mini_stream_size else 0

        regular_stream_sectors = 0
        for s in regular_streams:
            s.stream_size = len(s.data or b'')
            regular_stream_sectors += (s.stream_size + self.sector_size - 1) // self.sector_size

        total_data_sectors = dir_sectors_needed + mini_fat_sectors + mini_stream_sectors + regular_stream_sectors
        fat_sectors_needed = (total_data_sectors + entries_per_fat_sector - 1) // entries_per_fat_sector

        for _ in range(10):
            total_sectors = fat_sectors_needed + total_data_sectors
            new_fat_sectors = (total_sectors + entries_per_fat_sector - 1) // entries_per_fat_sector
            if new_fat_sectors == fat_sectors_needed:
                break
            fat_sectors_needed = new_fat_sectors

        # Keep at least 3 FAT sectors for Excel compatibility
        if fat_sectors_needed < 3:
            fat_sectors_needed = 3

        # Layout with FAT sector 0 and remaining FAT sectors at end
        current_sector = 0
        fat_sectors_list = [0]
        current_sector += 1

        layout['dir_start'] = current_sector
        layout['dir_sectors'] = dir_sectors_needed
        current_sector += dir_sectors_needed

        layout['mini_fat_start'] = current_sector if mini_fat_sectors else self.ENDOFCHAIN
        layout['mini_fat_sectors'] = mini_fat_sectors
        current_sector += mini_fat_sectors

        layout['mini_stream_start'] = current_sector if mini_stream_sectors else self.ENDOFCHAIN
        layout['mini_stream_sectors'] = mini_stream_sectors
        current_sector += mini_stream_sectors

        for s in regular_streams:
            sectors_needed = (s.stream_size + self.sector_size - 1) // self.sector_size
            if sectors_needed == 0:
                sectors_needed = 1
            s.starting_sector = current_sector
            current_sector += sectors_needed

        for _ in range(fat_sectors_needed - 1):
            fat_sectors_list.append(current_sector)
            current_sector += 1

        layout['stream_info'] = {
            'mini_streams': mini_streams,
            'regular_streams': regular_streams,
            'mini_stream_data': mini_stream_data
        }
        layout['total_sectors'] = current_sector
        layout['fat_sectors'] = fat_sectors_needed
        layout['fat_sectors_list'] = fat_sectors_list

        if mini_stream_sectors:
            self.root.starting_sector = layout['mini_stream_start']
        else:
            self.root.starting_sector = self.ENDOFCHAIN
            self.root.stream_size = 0

        return layout

    def _build_header(self, layout):
        header = io.BytesIO()

        header.write(struct.pack('<Q', self.HEADER_SIGNATURE))
        header.write(self.HEADER_CLSID)
        header.write(struct.pack('<H', self.MINOR_VERSION))
        header.write(struct.pack('<H', self.major_version))
        header.write(struct.pack('<H', self.BYTE_ORDER))
        header.write(struct.pack('<H', self.sector_shift))
        header.write(struct.pack('<H', 6))
        header.write(b'\x00' * 6)

        if self.major_version == self.MAJOR_VERSION_4:
            header.write(struct.pack('<I', layout['total_sectors']))
        else:
            header.write(struct.pack('<I', 0))

        header.write(struct.pack('<I', layout['fat_sectors']))
        header.write(struct.pack('<I', layout['dir_start']))
        header.write(struct.pack('<I', 0))
        header.write(struct.pack('<I', self.MINI_STREAM_CUTOFF))
        header.write(struct.pack('<I', layout['mini_fat_start']))
        header.write(struct.pack('<I', layout['mini_fat_sectors']))
        header.write(struct.pack('<I', self.ENDOFCHAIN))
        header.write(struct.pack('<I', 0))

        for i in range(109):
            if i < len(layout['fat_sectors_list']):
                header.write(struct.pack('<I', layout['fat_sectors_list'][i]))
            else:
                header.write(struct.pack('<I', self.FREESECT))

        data = header.getvalue()
        if len(data) < 512:
            data += b'\x00' * (512 - len(data))
        return data[:512]

    def _build_fat(self, layout):
        total_sectors = layout['total_sectors']
        fat = [self.FREESECT] * total_sectors

        # FAT sectors
        for sector in layout['fat_sectors_list']:
            fat[sector] = self.FATSECT

        # Directory sectors
        for i in range(layout['dir_sectors']):
            sector = layout['dir_start'] + i
            next_sector = layout['dir_start'] + i + 1
            fat[sector] = next_sector if i < layout['dir_sectors'] - 1 else self.ENDOFCHAIN

        # MiniFAT sectors
        if layout['mini_fat_sectors']:
            for i in range(layout['mini_fat_sectors']):
                sector = layout['mini_fat_start'] + i
                next_sector = layout['mini_fat_start'] + i + 1
                fat[sector] = next_sector if i < layout['mini_fat_sectors'] - 1 else self.ENDOFCHAIN

        # Mini stream sectors
        if layout['mini_stream_sectors']:
            for i in range(layout['mini_stream_sectors']):
                sector = layout['mini_stream_start'] + i
                next_sector = layout['mini_stream_start'] + i + 1
                fat[sector] = next_sector if i < layout['mini_stream_sectors'] - 1 else self.ENDOFCHAIN

        # Regular stream sectors
        for s in layout['stream_info']['regular_streams']:
            sectors_needed = (s.stream_size + self.sector_size - 1) // self.sector_size
            for i in range(sectors_needed):
                sector = s.starting_sector + i
                next_sector = s.starting_sector + i + 1
                fat[sector] = next_sector if i < sectors_needed - 1 else self.ENDOFCHAIN

        fat_bytes = io.BytesIO()
        for entry in fat:
            fat_bytes.write(struct.pack('<I', entry))

        entries_per_sector = self.sector_size // 4
        total_fat_entries = layout['fat_sectors'] * entries_per_sector
        entries_written = len(fat)
        for _ in range(entries_written, total_fat_entries):
            fat_bytes.write(struct.pack('<I', self.FREESECT))

        return fat_bytes.getvalue()

    def _build_minifat(self, layout):
        mini_streams = layout['stream_info']['mini_streams']
        mini_stream_data = layout['stream_info']['mini_stream_data']
        if not mini_stream_data:
            return b''

        mini_sector_count = len(mini_stream_data) // self.MINI_SECTOR_SIZE
        mini_fat = [self.FREESECT] * mini_sector_count

        for s in mini_streams:
            start = s.starting_sector
            sectors = (s.stream_size + self.MINI_SECTOR_SIZE - 1) // self.MINI_SECTOR_SIZE
            for i in range(sectors):
                idx = start + i
                mini_fat[idx] = (start + i + 1) if i < sectors - 1 else self.ENDOFCHAIN

        data = io.BytesIO()
        for entry in mini_fat:
            data.write(struct.pack('<I', entry))

        entries_per_sector = self.sector_size // 4
        total_entries = layout['mini_fat_sectors'] * entries_per_sector
        entries_written = len(mini_fat)
        for _ in range(entries_written, total_entries):
            data.write(struct.pack('<I', self.FREESECT))

        return data.getvalue()

    def _build_directory(self, nodes):
        directory = io.BytesIO()

        for node in nodes:
            left = node.left.did if node.left is not None else 0xFFFFFFFF
            right = node.right.did if node.right is not None else 0xFFFFFFFF
            child = node.child.did if node.child is not None else 0xFFFFFFFF

            starting_sector = node.starting_sector
            stream_size = node.stream_size
            if node.obj_type == self.STGTY_STORAGE:
                # Excel uses 0 for storage entries with no stream.
                starting_sector = 0
                stream_size = 0

            entry = self._create_directory_entry(
                name=node.name,
                obj_type=node.obj_type,
                color=node.color,
                left_sibling=left,
                right_sibling=right,
                child_did=child,
                clsid=b'\x00' * 16,
                state_bits=0,
                creation_time=0,
                modified_time=0,
                starting_sector=starting_sector,
                stream_size=stream_size
            )
            directory.write(entry)

        entries_per_sector = self.sector_size // 128
        total_entries = ((len(nodes) + entries_per_sector - 1) // entries_per_sector) * entries_per_sector
        for _ in range(len(nodes), total_entries):
            directory.write(b'\xFF' * 128)

        return directory.getvalue()

    def _create_directory_entry(self, name, obj_type, color, left_sibling, right_sibling,
                                child_did, clsid, state_bits, creation_time, modified_time,
                                starting_sector, stream_size):
        entry = io.BytesIO()

        name_bytes = name.encode('utf-16le')
        if len(name_bytes) > 62:
            name_bytes = name_bytes[:62]
        entry.write(name_bytes)
        entry.write(b'\x00' * (64 - len(name_bytes)))

        name_len = len(name_bytes) + 2
        entry.write(struct.pack('<H', name_len))
        entry.write(struct.pack('<B', obj_type))
        entry.write(struct.pack('<B', color))
        entry.write(struct.pack('<I', left_sibling))
        entry.write(struct.pack('<I', right_sibling))
        entry.write(struct.pack('<I', child_did))
        entry.write(clsid)
        entry.write(struct.pack('<I', state_bits))
        entry.write(struct.pack('<Q', creation_time))
        entry.write(struct.pack('<Q', modified_time))
        entry.write(struct.pack('<I', starting_sector))
        entry.write(struct.pack('<Q', stream_size))

        data = entry.getvalue()
        assert len(data) == 128
        return data

    def _build_stream_sectors(self, layout):
        sector_map = {}

        mini_stream_data = layout['stream_info']['mini_stream_data']
        if mini_stream_data:
            offset = 0
            for i in range(layout['mini_stream_sectors']):
                sector = mini_stream_data[offset:offset + self.sector_size]
                if len(sector) < self.sector_size:
                    sector += b'\x00' * (self.sector_size - len(sector))
                sector_map[layout['mini_stream_start'] + i] = sector
                offset += self.sector_size

        for s in layout['stream_info']['regular_streams']:
            data = s.data or b''
            offset = 0
            sectors_needed = (s.stream_size + self.sector_size - 1) // self.sector_size
            for i in range(sectors_needed):
                sector = data[offset:offset + self.sector_size]
                if len(sector) < self.sector_size:
                    sector += b'\x00' * (self.sector_size - len(sector))
                sector_map[s.starting_sector + i] = sector
                offset += self.sector_size

        return sector_map

    def _split_into_sectors(self, data):
        sectors = []
        offset = 0
        while offset < len(data):
            sector = data[offset:offset + self.sector_size]
            if len(sector) < self.sector_size:
                sector += b'\x00' * (self.sector_size - len(sector))
            sectors.append(sector)
            offset += self.sector_size
        return sectors
