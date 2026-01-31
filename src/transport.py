import threading
import time
import serial

# Transport constants
# Bits per character on the serial line (start + data + parity + stop)
BITS_PER_CHAR = 10
# Number of characters to wait for the first response byte (was requested as 20)
START_CHARS = 200
# Minimum start timeout in seconds (prevents extremely small timeouts at high baudrates)
MIN_START_TIMEOUT = 0.5


def compute_crc(data: bytes) -> int:
    # Follow Modbus RTU CRC algorithm exactly, masking to 16 bits as we go.
    crc = 0xFFFF
    for a in data:
        crc ^= (a & 0xFF)
        for _ in range(8):
            lsb = crc & 0x0001
            crc >>= 1
            if lsb:
                crc ^= 0xA001
            crc &= 0xFFFF
    return crc & 0xFFFF


class SerialTransport:
    def __init__(self):
        self.ser = None
        self.lock = threading.Lock()

    def open(self, port, baudrate, parity):
        if self.ser and self.ser.is_open:
            self.close()
        p = {'N': serial.PARITY_NONE, 'E': serial.PARITY_EVEN, 'O': serial.PARITY_ODD}[parity]
        self.ser = serial.Serial(port=port, baudrate=baudrate, bytesize=8, parity=p, stopbits=1, timeout=0)

    def close(self):
        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None

    def char_time(self):
        if not self.ser:
            return 0.001
        parity = self.ser.parity
        parity_bit = 1 if parity != serial.PARITY_NONE else 0
        bits = 1 + 8 + parity_bit + 1
        return bits / float(self.ser.baudrate)

    def send_and_receive(self, request: bytes, expect_response=True):
        if not self.ser or not self.ser.is_open:
            raise IOError('Serial port not open')

        with self.lock:
            try:
                self.ser.reset_input_buffer()
            except Exception:
                pass
            self.ser.write(request)
            self.ser.flush()

            if not expect_response:
                return b''

            # Wait-for-first-byte timeout: use formula 1/baud * START_CHARS * BITS_PER_CHAR
            calc = (1.0 / float(self.ser.baudrate)) * float(START_CHARS) * float(BITS_PER_CHAR)
            # enforce a sensible minimum timeout to avoid overly short waits at high baudrates
            start_timeout = max(calc, MIN_START_TIMEOUT)
            start = time.time()
            first = b''
            while True:
                if time.time() - start > start_timeout:
                    raise TimeoutError('No response started within timeout')
                b = self.ser.read(1)
                if b:
                    first = b
                    break
                time.sleep(0.001)

            header = first + self._read_exact(2)
            if len(header) < 3:
                raise IOError('Incomplete header')
            cmd = header[1]
            if cmd == 0x03:
                byte_num = header[2]
                data = self._read_exact(byte_num + 2)
                resp = header + data
                if not self._check_crc(resp):
                    raise IOError('CRC mismatch')
                return resp
            elif cmd == 0x04:
                # function 0x04 (read input registers) has same structure as 0x03
                byte_num = header[2]
                data = self._read_exact(byte_num + 2)
                resp = header + data
                if not self._check_crc(resp):
                    raise IOError('CRC mismatch')
                return resp
            elif cmd == 0x06:
                rest = self._read_exact(4 + 2)
                resp = header + rest
                if not self._check_crc(resp):
                    raise IOError('CRC mismatch')
                return resp
            else:
                buf = b''
                t0 = time.time()
                while True:
                    chunk = self.ser.read(256)
                    if chunk:
                        buf += chunk
                        t0 = time.time()
                    else:
                        if time.time() - t0 > 0.05:
                            break
                        time.sleep(0.001)
                resp = header + buf
                if len(resp) >= 3 and self._check_crc(resp):
                    return resp
                else:
                    raise IOError('Unknown or invalid response')

    def _read_exact(self, n, timeout=1.0):
        data = b''
        start = time.time()
        while len(data) < n:
            chunk = self.ser.read(n - len(data))
            if chunk:
                data += chunk
            else:
                if time.time() - start > timeout:
                    break
                time.sleep(0.001)
        return data

    def _check_crc(self, data: bytes) -> bool:
        if len(data) < 3:
            return False
        crc_recv = data[-2] | (data[-1] << 8)
        calc = compute_crc(data[:-2])
        return crc_recv == calc

    def read_status(self, drive_id: int, start_addr: int, count: int = 1, func: int = 0x04):
        """Read `count` 16-bit status registers starting at `start_addr` using function `func` (0x04 or 0x03).

        Returns a list of integers of length `count` on success.
        Raises exceptions on transport errors.
        """
        if func not in (0x03, 0x04):
            raise ValueError('Unsupported function for status read')
        # build request: Addr(1) FC(1) START_H START_L NUM_H NUM_L CRC_L CRC_H
        try:
            import struct as _struct
            req = _struct.pack('>B B H H', drive_id, func, start_addr, count)
            crc = compute_crc(req)
            req += _struct.pack('<H', crc)
            resp = self.send_and_receive(req)
            # resp[2] = byte count, then two-byte values
            if len(resp) < 3:
                raise IOError('Short response')
            byte_num = resp[2]
            data = resp[3:3+byte_num]
            vals = []
            for i in range(0, len(data), 2):
                hi = data[i]
                lo = data[i+1] if i+1 < len(data) else 0
                vals.append((hi << 8) | lo)
            # ensure list length matches requested count (pad with zeros if missing)
            while len(vals) < count:
                vals.append(0)
            return vals[:count]
        except Exception:
            raise
