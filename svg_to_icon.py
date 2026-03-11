from PIL import Image
import cairosvg

# 1. SVG转256x256 PNG（基础尺寸）
cairosvg.svg2png(url="icon.svg", write_to="icon_temp.png", output_width=256, output_height=256)
# 2. 生成多尺寸ICO（覆盖Windows所有场景）
img = Image.open("icon_temp.png")
# 必须包含这些尺寸：16x16(任务栏)、32x32(小图标)、48x48(中等图标)、256x256(高清)
img.save(
    "icon.ico",
    format="ICO",
    sizes=[(16,16), (32,32), (48,48), (256,256)]
)
# 3. 删除临时文件
import os
os.remove("icon_temp.png")
print("标准多尺寸ICO生成成功！")
