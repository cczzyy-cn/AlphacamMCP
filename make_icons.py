"""生成 32x32 BMP 图标 - 剪刀 + 箭头"""
import struct, os

def make_bmp(filename, pixels_32x32):
    """32x32 24-bit BMP"""
    width, height = 32, 32
    row_size = ((width * 3 + 3) // 4) * 4
    data_size = row_size * height
    file_size = 54 + data_size
    
    buf = bytearray()
    buf.extend(struct.pack('<H', 0x4D42))  # BM
    buf.extend(struct.pack('<I', file_size))
    buf.extend(struct.pack('<HH', 0, 0))
    buf.extend(struct.pack('<I', 54))
    buf.extend(struct.pack('<I', 40))
    buf.extend(struct.pack('<i', width))
    buf.extend(struct.pack('<i', height))
    buf.extend(struct.pack('<H', 1))
    buf.extend(struct.pack('<H', 24))
    buf.extend(struct.pack('<I', 0))
    buf.extend(struct.pack('<I', data_size))
    buf.extend(struct.pack('<i', 2835))
    buf.extend(struct.pack('<i', 2835))
    buf.extend(struct.pack('<I', 0))
    buf.extend(struct.pack('<I', 0))
    
    for y in range(height - 1, -1, -1):
        row = bytearray()
        for x in range(width):
            r, g, b = pixels_32x32[y][x]
            row.extend(struct.pack('BBB', b, g, r))
        row.extend(b'\x00' * (row_size - width * 3))
        buf.extend(row)
    
    with open(filename, 'wb') as f:
        f.write(buf)
    print(f"  {filename} - {os.path.getsize(filename)} bytes")

W = (255,255,255)
B = (0,0,0)
G = (180,180,180)
R = (200,0,0)

# === 剪刀图标 32x32 ===
s = [[W]*32 for _ in range(32)]
# 剪刀主体 - 两个椭圆环
for y in range(32):
    for x in range(32):
        # 左侧刀片环 (中心 8,12)
        dx1, dy1 = x-8, y-12
        d1 = (dx1*dx1 + dy1*dy1*3)**0.5
        # 右侧刀片环 (中心 22,12)  
        dx2, dy2 = x-22, y-12
        d2 = (dx2*dx2 + dy2*dy2*3)**0.5
        # 刀片线条
        if abs(d1-6) < 1.2 or abs(d2-6) < 1.2:
            s[y][x] = B
        # 刀柄
        if 10 <= x <= 20 and 24 <= y <= 28 and (x-15)*(x-15)+(y-26)*(y-26) < 6:
            s[y][x] = B
        if 12 <= x <= 18 and 22 <= y <= 25 and (x-15) > (y-23)+2:
            s[y][x] = B
        # 螺丝
        if (x-15)*(x-15)+(y-12)*(y-12) < 2.5:
            s[y][x] = G
s[10][8] = B; s[10][9] = B; s[11][10] = B
s[10][21] = B; s[10][22] = B; s[11][20] = B
# 补一些像素让剪刀更明显
for y in range(32):
    for x in range(32):
        if (x-6)*(x-6)+(y-18)*(y-18) < 2:
            s[y][x] = B
        if (x-24)*(x-24)+(y-18)*(y-18) < 2:
            s[y][x] = B

# === 箭头偏移图标 32x32 ===
a = [[W]*32 for _ in range(32)]

# 向右箭头
for y in range(32):
    for x in range(32):
        # 箭头杆 (水平条)
        if 4 <= y <= 12 and 8 <= x <= 26:
            a[y][x] = B
        # 箭头头 (三角形)
        dx = x - 26
        dy = y - 8
        if dx >= 0 and abs(dy) <= dx and dx <= 6:
            a[y][x] = B

# 箭尾竖线 + 第二个箭头
for y in range(32):
    for x in range(32):
        # 第二个箭头向下
        if 16 <= x <= 24 and 14 <= y <= 28:
            a[y][x] = B
        # 箭头头向下
        dy2 = y - 28
        dx2 = x - 20
        if dy2 >= 0 and abs(dx2) <= dy2 and dy2 <= 6:
            a[y][x] = B

# 补X轴标记
for y in range(32):
    for x in range(32):
        if abs(x-27)+abs(y-8) < 1.5:
            a[y][x] = R
        if abs(x-20)+abs(y-29) < 1.5:
            a[y][x] = R

out = r'C:\Users\C\AppData\Roaming\reasonix\global-workspace\.reasonix\skills\alphacam-bridge'
make_bmp(f'{out}\\cut.bmp', s)
make_bmp(f'{out}\\offset.bmp', a)
