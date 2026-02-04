"""
Test XLSX Encryption Feature

This test verifies that XLSX files can be encrypted and decrypted correctly
according to ECMA-376 Part 2 specification.
"""

import unittest
import os
import sys

# Add parent directory to path to import aspose_cells
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from aspose_cells import (
    Workbook,
    AgileEncryptionParameters,
    CipherAlgorithm,
    HashAlgorithm,
    encrypt_xlsx,
    decrypt_xlsx
)


class TestEncryption(unittest.TestCase):
    """Test cases for XLSX encryption feature."""

    def setUp(self):
        """Set up test fixtures."""
        self.test_password = "Test Password 123!"
        self.wrong_password = "Wrong Password"
        self.output_dir = os.path.join(os.path.dirname(__file__), "outputfiles")
        os.makedirs(self.output_dir, exist_ok=True)

    def create_test_workbook(self):
        """Create a test workbook with various data."""
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "EncryptionTest"

        # Add various types of data
        ws.cells['A1'].value = "Text Data"
        ws.cells['B1'].value = 12345
        ws.cells['C1'].value = 3.14159
        ws.cells['D1'].value = "=B1*C1"  # Formula

        # Add styled cells
        ws.cells['A2'].value = "Styled Cell"
        ws.cells['A2'].style.font.bold = True
        ws.cells['A2'].style.font.color = "FF0000"

        # Add more data
        for i in range(10):
            ws.cells[f'A{i+3}'].value = f"Row {i+1}"
            ws.cells[f'B{i+3}'].value = i * 100

        return wb

    def test_basic_encryption_decryption(self):
        """Test basic encryption and decryption with default settings."""
        print("\n" + "="*70)
        print("Test: Basic Encryption and Decryption")
        print("="*70)

        # Create test workbook
        print("\nCreating test workbook...")
        wb = self.create_test_workbook()

        # Save with encryption
        encrypted_file = os.path.join(self.output_dir, "test_encrypted_basic.xlsx")
        print(f"Saving encrypted file to {encrypted_file}...")
        wb.save(encrypted_file, password=self.test_password)
        self.assertTrue(os.path.exists(encrypted_file))
        print(f"  [OK] Encrypted file created ({os.path.getsize(encrypted_file)} bytes)")

        # Verify file is encrypted (CFB format)
        with open(encrypted_file, 'rb') as f:
            header = f.read(8)
            # CFB signature: D0 CF 11 E0 A1 B1 1A E1
            self.assertEqual(header, b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
            print("  [OK] File is in CFB format (encrypted)")

        # Decrypt and load
        print("\nDecrypting and loading workbook...")
        wb_loaded = Workbook(encrypted_file, password=self.test_password)
        ws_loaded = wb_loaded.worksheets[0]

        # Verify data
        self.assertEqual(ws_loaded.name, "EncryptionTest")
        self.assertEqual(ws_loaded.cells['A1'].value, "Text Data")
        self.assertEqual(ws_loaded.cells['B1'].value, 12345)
        self.assertAlmostEqual(float(ws_loaded.cells['C1'].value), 3.14159, places=5)
        print("  [OK] Data integrity verified")

        # Verify styled cell
        self.assertEqual(ws_loaded.cells['A2'].value, "Styled Cell")
        self.assertTrue(ws_loaded.cells['A2'].style.font.bold)
        print("  [OK] Formatting preserved")

        print("\n[OK] Basic encryption/decryption test passed!")
        print("="*70 + "\n")

    def test_wrong_password(self):
        """Test that wrong password is rejected."""
        print("\n" + "="*70)
        print("Test: Wrong Password Rejection")
        print("="*70)

        # Create and encrypt workbook
        wb = self.create_test_workbook()
        encrypted_file = os.path.join(self.output_dir, "test_encrypted_wrong_pwd.xlsx")
        wb.save(encrypted_file, password=self.test_password)
        print(f"\nEncrypted file created with password: '{self.test_password}'")

        # Try to load with wrong password
        print(f"Attempting to load with wrong password: '{self.wrong_password}'...")
        with self.assertRaises(ValueError) as context:
            wb_loaded = Workbook(encrypted_file, password=self.wrong_password)

        print(f"  [OK] Correctly rejected with error: {context.exception}")

        # Verify correct password still works
        print(f"\nVerifying correct password still works...")
        wb_loaded = Workbook(encrypted_file, password=self.test_password)
        self.assertEqual(wb_loaded.worksheets[0].cells['A1'].value, "Text Data")
        print("  [OK] Correct password accepted")

        print("\n[OK] Wrong password rejection test passed!")
        print("="*70 + "\n")

    def test_no_password_on_encrypted_file(self):
        """Test that encrypted file without password is rejected."""
        print("\n" + "="*70)
        print("Test: No Password on Encrypted File")
        print("="*70)

        # Create and encrypt workbook
        wb = self.create_test_workbook()
        encrypted_file = os.path.join(self.output_dir, "test_encrypted_no_pwd.xlsx")
        wb.save(encrypted_file, password=self.test_password)
        print(f"\nEncrypted file created")

        # Try to load without password
        print("Attempting to load encrypted file without password...")
        with self.assertRaises(ValueError) as context:
            wb_loaded = Workbook(encrypted_file)

        self.assertIn("encrypted", str(context.exception).lower())
        print(f"  [OK] Correctly rejected with error: {context.exception}")

        print("\n[OK] No password rejection test passed!")
        print("="*70 + "\n")

    def test_aes_256_sha512(self):
        """Test AES-256 with SHA-512 (Office 2013+ default)."""
        print("\n" + "="*70)
        print("Test: AES-256 + SHA-512 Encryption")
        print("="*70)

        wb = self.create_test_workbook()

        # Create custom encryption parameters
        params = AgileEncryptionParameters(
            cipher_algorithm=CipherAlgorithm.AES_256,
            hash_algorithm=HashAlgorithm.SHA512,
            spin_count=100000
        )
        print("\nEncryption parameters:")
        print(f"  Cipher: AES-256 ({params.cipher_algorithm.key_bits} bits)")
        print(f"  Hash: SHA-512 ({params.hash_algorithm.hash_bytes} bytes)")
        print(f"  Spin count: {params.spin_count:,}")

        # Save with custom parameters
        encrypted_file = os.path.join(self.output_dir, "test_encrypted_aes256_sha512.xlsx")
        print(f"\nSaving with AES-256/SHA-512...")
        wb.save(encrypted_file, password=self.test_password, encryption_params=params)
        self.assertTrue(os.path.exists(encrypted_file))
        print(f"  [OK] Encrypted file created")

        # Load and verify
        print("\nDecrypting and verifying...")
        wb_loaded = Workbook(encrypted_file, password=self.test_password)
        self.assertEqual(wb_loaded.worksheets[0].cells['A1'].value, "Text Data")
        print("  [OK] Data verified")

        print("\n[OK] AES-256/SHA-512 encryption test passed!")
        print("="*70 + "\n")

    def test_aes_128_sha1(self):
        """Test AES-128 with SHA-1 (Office 2010 default)."""
        print("\n" + "="*70)
        print("Test: AES-128 + SHA-1 Encryption")
        print("="*70)

        wb = self.create_test_workbook()

        # Create encryption parameters for Office 2010 compatibility
        params = AgileEncryptionParameters(
            cipher_algorithm=CipherAlgorithm.AES_128,
            hash_algorithm=HashAlgorithm.SHA1,
            spin_count=100000
        )
        print("\nEncryption parameters:")
        print(f"  Cipher: AES-128 ({params.cipher_algorithm.key_bits} bits)")
        print(f"  Hash: SHA-1 ({params.hash_algorithm.hash_bytes} bytes)")
        print(f"  Spin count: {params.spin_count:,}")

        # Save with custom parameters
        encrypted_file = os.path.join(self.output_dir, "test_encrypted_aes128_sha1.xlsx")
        print(f"\nSaving with AES-128/SHA-1...")
        wb.save(encrypted_file, password=self.test_password, encryption_params=params)
        self.assertTrue(os.path.exists(encrypted_file))
        print(f"  [OK] Encrypted file created")

        # Load and verify
        print("\nDecrypting and verifying...")
        wb_loaded = Workbook(encrypted_file, password=self.test_password)
        self.assertEqual(wb_loaded.worksheets[0].cells['A1'].value, "Text Data")
        print("  [OK] Data verified")

        print("\n[OK] AES-128/SHA-1 encryption test passed!")
        print("="*70 + "\n")

    def test_different_hash_algorithms(self):
        """Test different hash algorithms."""
        print("\n" + "="*70)
        print("Test: Different Hash Algorithms")
        print("="*70)

        hash_algorithms = [
            (HashAlgorithm.SHA1, "SHA1"),
            (HashAlgorithm.SHA256, "SHA256"),
            (HashAlgorithm.SHA384, "SHA384"),
            (HashAlgorithm.SHA512, "SHA512")
        ]

        print("\nTesting hash algorithms:")
        for hash_alg, name in hash_algorithms:
            print(f"\n  Testing {name}...")
            wb = self.create_test_workbook()

            params = AgileEncryptionParameters(
                cipher_algorithm=CipherAlgorithm.AES_256,
                hash_algorithm=hash_alg,
                spin_count=50000  # Reduced for faster testing
            )

            encrypted_file = os.path.join(self.output_dir, f"test_encrypted_{name.lower()}.xlsx")
            wb.save(encrypted_file, password=self.test_password, encryption_params=params)

            # Verify roundtrip
            wb_loaded = Workbook(encrypted_file, password=self.test_password)
            self.assertEqual(wb_loaded.worksheets[0].cells['A1'].value, "Text Data")
            print(f"    [OK] {name} encryption/decryption successful")

        print("\n[OK] All hash algorithms test passed!")
        print("="*70 + "\n")

    def test_utility_functions(self):
        """Test utility functions (encrypt_xlsx, decrypt_xlsx)."""
        print("\n" + "="*70)
        print("Test: Utility Functions")
        print("="*70)

        # Create unencrypted file
        wb = self.create_test_workbook()
        unencrypted_file = os.path.join(self.output_dir, "test_unencrypted.xlsx")
        wb.save(unencrypted_file)
        print(f"\nCreated unencrypted file: {unencrypted_file}")

        # Encrypt using utility function
        encrypted_file = os.path.join(self.output_dir, "test_util_encrypted.xlsx")
        print(f"\nEncrypting with encrypt_xlsx()...")
        encrypt_xlsx(unencrypted_file, encrypted_file, self.test_password)
        self.assertTrue(os.path.exists(encrypted_file))
        print(f"  [OK] Encrypted file created")

        # Decrypt using utility function
        decrypted_file = os.path.join(self.output_dir, "test_util_decrypted.xlsx")
        print(f"\nDecrypting with decrypt_xlsx()...")
        decrypt_xlsx(encrypted_file, decrypted_file, self.test_password)
        self.assertTrue(os.path.exists(decrypted_file))
        print(f"  [OK] Decrypted file created")

        # Verify decrypted file
        wb_loaded = Workbook(decrypted_file)
        self.assertEqual(wb_loaded.worksheets[0].cells['A1'].value, "Text Data")
        print("  [OK] Data verified")

        print("\n[OK] Utility functions test passed!")
        print("="*70 + "\n")

    def test_comprehensive_roundtrip(self):
        """Test comprehensive roundtrip with complex workbook."""
        print("\n" + "="*70)
        print("Test: Comprehensive Roundtrip")
        print("="*70)

        # Create complex workbook
        print("\nCreating complex workbook...")
        wb = Workbook()

        # First worksheet with various data
        ws1 = wb.worksheets[0]
        ws1.name = "DataSheet"
        ws1.cells['A1'].value = "Product"
        ws1.cells['B1'].value = "Price"
        ws1.cells['C1'].value = "Quantity"
        ws1.cells['D1'].value = "Total"

        products = [("Widget", 10.50, 100), ("Gadget", 25.00, 50), ("Gizmo", 5.75, 200)]
        for i, (product, price, qty) in enumerate(products, start=2):
            ws1.cells[f'A{i}'].value = product
            ws1.cells[f'B{i}'].value = price
            ws1.cells[f'C{i}'].value = qty
            ws1.cells[f'D{i}'].value = f"=B{i}*C{i}"

        # Second worksheet
        from aspose_cells import Worksheet
        ws2 = Worksheet("Summary")
        wb.worksheets.append(ws2)
        ws2.cells['A1'].value = "Total Products"
        ws2.cells['B1'].value = len(products)

        print(f"  Created workbook with {len(wb.worksheets)} sheets")
        print(f"  Sheet 1: '{ws1.name}' ({products.__len__()} products)")
        print(f"  Sheet 2: '{ws2.name}'")

        # Save with encryption
        encrypted_file = os.path.join(self.output_dir, "test_encrypted_comprehensive.xlsx")
        print(f"\nEncrypting comprehensive workbook...")
        wb.save(encrypted_file, password=self.test_password)
        file_size = os.path.getsize(encrypted_file)
        print(f"  [OK] Encrypted file created ({file_size} bytes)")

        # Load and verify
        print("\nDecrypting and verifying comprehensive workbook...")
        wb_loaded = Workbook(encrypted_file, password=self.test_password)

        # Verify structure
        self.assertEqual(len(wb_loaded.worksheets), 2)
        self.assertEqual(wb_loaded.worksheets[0].name, "DataSheet")
        self.assertEqual(wb_loaded.worksheets[1].name, "Summary")
        print("  [OK] Workbook structure verified")

        # Verify data
        ws1_loaded = wb_loaded.worksheets[0]
        self.assertEqual(ws1_loaded.cells['A1'].value, "Product")
        self.assertEqual(ws1_loaded.cells['A2'].value, "Widget")
        self.assertAlmostEqual(float(ws1_loaded.cells['B2'].value), 10.50, places=2)
        print("  [OK] Data Sheet verified")

        ws2_loaded = wb_loaded.worksheets[1]
        self.assertEqual(ws2_loaded.cells['A1'].value, "Total Products")
        self.assertEqual(ws2_loaded.cells['B1'].value, len(products))
        print("  [OK] Summary Sheet verified")

        print("\n[OK] Comprehensive roundtrip test passed!")
        print("="*70 + "\n")


    def test_simple_encryption_with_hello_password(self):
        """Test simple encryption with password 'hello' - populate data and save."""
        print("\n" + "="*70)
        print("Test: Simple Encryption with Password 'hello'")
        print("="*70)

        # Create a new workbook
        print("\nCreating workbook...")
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.name = "TestData"

        # Populate some data
        print("Populating data...")
        ws.cells['A1'].value = "Name"
        ws.cells['B1'].value = "Age"
        ws.cells['C1'].value = "Score"

        ws.cells['A2'].value = "Alice"
        ws.cells['B2'].value = 25
        ws.cells['C2'].value = 95.5

        ws.cells['A3'].value = "Bob"
        ws.cells['B3'].value = 30
        ws.cells['C3'].value = 87.0

        ws.cells['A4'].value = "Charlie"
        ws.cells['B4'].value = 28
        ws.cells['C4'].value = 92.3

        print("  [OK] Data populated:")
        print("    - A1: 'Name', B1: 'Age', C1: 'Score'")
        print("    - A2: 'Alice', B2: 25, C2: 95.5")
        print("    - A3: 'Bob', B3: 30, C3: 87.0")
        print("    - A4: 'Charlie', B4: 28, C4: 92.3")

        # Save to non-encrypted XLSX file first
        unencrypted_file = os.path.join(self.output_dir, "test_unencrypted.xlsx")
        print(f"\nSaving to non-encrypted file: {unencrypted_file}")
        wb.save(unencrypted_file)
        
        self.assertTrue(os.path.exists(unencrypted_file))
        unencrypted_size = os.path.getsize(unencrypted_file)
        print(f"  [OK] Non-encrypted file created ({unencrypted_size} bytes)")

        # Verify non-encrypted file is ZIP format (XLSX)
        with open(unencrypted_file, 'rb') as f:
            header = f.read(4)
            # ZIP signature: 50 4B 03 04 (PK..)
            self.assertEqual(header, b'\x50\x4b\x03\x04')
            print("  [OK] Non-encrypted file is in ZIP format (XLSX)")

        # Save to encrypted XLSX file with password "hello"
        password = "hello"
        encrypted_file = os.path.join(self.output_dir, "test_hello_password.xlsx")
        print(f"\nSaving to encrypted file: {encrypted_file}")
        print(f"Using password: '{password}'")
        wb.save(encrypted_file, password=password)
        
        self.assertTrue(os.path.exists(encrypted_file))
        encrypted_size = os.path.getsize(encrypted_file)
        print(f"  [OK] Encrypted file created ({encrypted_size} bytes)")
        print(f"  [INFO] Size comparison: {unencrypted_size} bytes (unencrypted) vs {encrypted_size} bytes (encrypted)")

        # Verify encrypted file is CFB format
        with open(encrypted_file, 'rb') as f:
            header = f.read(8)
            # CFB signature: D0 CF 11 E0 A1 B1 1A E1
            self.assertEqual(header, b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
            print("  [OK] Encrypted file is in CFB format")

        # Load and verify data with correct password
        print("\nLoading encrypted file with password 'hello'...")
        wb_loaded = Workbook(encrypted_file, password=password)
        ws_loaded = wb_loaded.worksheets[0]

        # Verify worksheet name
        self.assertEqual(ws_loaded.name, "TestData")
        print("  [OK] Worksheet name verified: 'TestData'")

        # Verify header row
        self.assertEqual(ws_loaded.cells['A1'].value, "Name")
        self.assertEqual(ws_loaded.cells['B1'].value, "Age")
        self.assertEqual(ws_loaded.cells['C1'].value, "Score")
        print("  [OK] Header row verified")

        # Verify data rows
        self.assertEqual(ws_loaded.cells['A2'].value, "Alice")
        self.assertEqual(ws_loaded.cells['B2'].value, 25)
        self.assertAlmostEqual(float(ws_loaded.cells['C2'].value), 95.5, places=1)
        print("  [OK] Row 2 verified: Alice, 25, 95.5")

        self.assertEqual(ws_loaded.cells['A3'].value, "Bob")
        self.assertEqual(ws_loaded.cells['B3'].value, 30)
        self.assertAlmostEqual(float(ws_loaded.cells['C3'].value), 87.0, places=1)
        print("  [OK] Row 3 verified: Bob, 30, 87.0")

        self.assertEqual(ws_loaded.cells['A4'].value, "Charlie")
        self.assertEqual(ws_loaded.cells['B4'].value, 28)
        self.assertAlmostEqual(float(ws_loaded.cells['C4'].value), 92.3, places=1)
        print("  [OK] Row 4 verified: Charlie, 28, 92.3")

        print("\n[OK] Simple encryption test with password 'hello' passed!")
        print("="*70 + "\n")


if __name__ == '__main__':
    # Create output directory if it doesn't exist
    os.makedirs('outputfiles', exist_ok=True)

    # Run tests
    unittest.main(verbosity=2)
