"""
Excel Password Hash Algorithm Implementation

This script implements the exact Excel password hashing algorithm for worksheet protection
as specified in ECMA-376 standard.
"""

def hash_password(password):
    """
    Hash a password for Excel worksheet protection using Excel's exact algorithm.

    Excel uses a legacy XOR-based hash for worksheet protection (not secure, but
    compatible with Excel). The algorithm processes characters from last to first.

    Args:
        password (str): The password to hash.

    Returns:
        str: Hex-encoded password hash for Excel compatibility.
    """
    if password is None:
        return None

    # Initialize hash value with password length
    hash_value = 0

    # Process characters from LAST to FIRST
    for char in reversed(password):
        # Get character code
        char_code = ord(char)

        # XOR with current hash
        hash_value ^= char_code

        # Rotate left by 1 bit (15 bits total since we use 16-bit values)
        hash_value = ((hash_value << 1) | (hash_value >> 14)) & 0x7FFF

    # XOR with password length
    hash_value ^= len(password)

    # XOR with constant 0xCE4B
    hash_value ^= 0xCE4B

    # Return as hex string (uppercase, 4 digits)
    return f"{hash_value:04X}"

# Test with password "abc"
password = "abc"
print(f"Testing password: '{password}'")
print(f"Expected (openpyxl): CC1A")
print()

print(f"Result: {hash_password(password)}")

# Manual trace
password_bytes = password.encode('utf-16-le')
print(f"UTF-16LE bytes: {[hex(b) for b in password_bytes[:6]]}")

hash_value = 0
for i in range(len(password_bytes) // 2):
    char = password_bytes[i*2] | (password_bytes[i*2+1] << 8)
    print(f"Char {i+1}: {hex(char)}")
    hash_value ^= char
    print(f"  After XOR: {hex(hash_value)}")
    hash_value = ((hash_value >> 1) | (hash_value << 15)) & 0xFFFF
    print(f"  After rotate: {hex(hash_value)}")
    hash_value = (hash_value + i) & 0xFFFF
    print(f"  After add index: {hex(hash_value)}")
