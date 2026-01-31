from transport import compute_crc

# Example from user: ADR=0x01, CMD=0x03, Start=0x0005 (they used example with 0x00 0x00 0x00 0x01 earlier),
# The user example: "read the parameter 0 segment 05 of the servo drive with station number 01H"
# They then show frame (ADR 01, CMD 03, STADDR 00 00, RNUM 00 01) but later example shows 00 01 00 02 etc.
# We'll test with the exact example the user gave at the end of their message:
# ADR 0x01, CMD 0x03, START 0x0005, NUM 0x0002 -> bytes: 01 03 00 05 00 02
msg = bytes([0x01, 0x03, 0x00, 0x05, 0x00, 0x02])
crc = compute_crc(msg)
low = crc & 0xFF
high = (crc >> 8) & 0xFF
print(f"Message: {msg.hex()} CRC=0x{crc:04X} low=0x{low:02X} high=0x{high:02X}")
# Also test the other example the user mentioned with CRC=0x3794 expectation.
# If we compute for bytes 01 03 00 01 00 02 (start 0x0001, num 0x0002) we expect CRC 0x3794 per user's note.
msg2 = bytes([0x01,0x03,0x00,0x01,0x00,0x02])
crc2 = compute_crc(msg2)
print(f"Message2: {msg2.hex()} CRC=0x{crc2:04X} low=0x{crc2 & 0xFF:02X} high=0x{(crc2>>8)&0xFF:02X}")
