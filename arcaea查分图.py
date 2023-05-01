from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
import datetime
import cv2
import numpy as np
import sys


# 文件路径
if getattr(sys, 'frozen', False):
    current_file_path = sys.executable
else:
    current_file_path = __file__
current_file_dir = os.path.dirname(current_file_path)

# 设置图片宽高和每个方格的大小
img_width = 1280
img_height = 2048
list_width = 1200
list_height = 1600
box_width = list_width // 4
box_height = list_height // 8

# 读取Excel文件
excel_file = pd.ExcelFile(os.path.join(current_file_dir, 'Arcaea查分表.xlsx'))
info = excel_file.parse('b30')

# 加载字体和图片
font_path = os.path.join(current_file_dir + '/fonts', 'Kazesawa-Regular.ttf')
font = ImageFont.truetype(font_path, size=36)
score_font_path = os.path.join(current_file_dir + '/fonts', 'GeosansLight.ttf')
score_font = ImageFont.truetype(score_font_path, size=36)
ptt_font_path = os.path.join(current_file_dir + '/fonts', 'Exo-Medium.ttf')
ptt_font = ImageFont.truetype(ptt_font_path, size=30)
name_font_path = os.path.join(current_file_dir + '/fonts', 'Kazesawa-Light.ttf')
name_font = ImageFont.truetype(name_font_path, size=32)

bg_path = os.path.join(current_file_dir + '/pictures', 'background.jpg')

# 缩短曲名（大写字母算作两个小写字母宽度）
def count_lowercase(s):
    count = 0
    for c in s:
        if ord('A') <= ord(c) <= ord('Z'):
            count += 2
        elif ord('a') <= ord(c) <= ord('z'):
            count += 1
    return count
def shorten_string(s):
    if count_lowercase(s) > 18:
        s = s[:13] + "..."
    return s

# 绘制透明方格
def drawRect(img, pos, **kwargs):
    transp = Image.new('RGBA', img.size, (0,0,0,0))
    draw = ImageDraw.Draw(transp, "RGBA")
    draw.rectangle(pos, **kwargs)
    img.paste(Image.alpha_composite(img, transp))

# 右上表头
ptt = float(info.iloc[0, 9])
if(ptt >= 12.5):
    ptt_pic_path = os.path.join(current_file_dir + '/pictures', 'rating_6.png')
elif(ptt >= 12):
    ptt_pic_path = os.path.join(current_file_dir + '/pictures', 'rating_5.png')
elif(ptt >= 11):
    ptt_pic_path = os.path.join(current_file_dir + '/pictures', 'rating_4.png')
elif(ptt >= 10):
    ptt_pic_path = os.path.join(current_file_dir + '/pictures', 'rating_3.png')
elif(ptt >= 7):
    ptt_pic_path = os.path.join(current_file_dir + '/pictures', 'rating_2.png')
else:
    ptt_pic_path = os.path.join(current_file_dir + '/pictures', 'rating_1.png')

def add_alpha_channel(img):
    """ 为jpg图像添加alpha通道 """

    b_channel, g_channel, r_channel = cv2.split(img) # 剥离jpg图像通道
    alpha_channel = np.ones(b_channel.shape, dtype=b_channel.dtype) * 255 # 创建Alpha通道

    img_new = cv2.merge((b_channel, g_channel, r_channel, alpha_channel)) # 融合通道
    return img_new

def merge_img(jpg_img, png_img, y1, y2, x1, x2):
    """ 将png透明图像与jpg图像叠加 
        y1,y2,x1,x2为叠加位置坐标值
    """
    
    # 判断jpg图像是否已经为4通道
    if jpg_img.shape[2] == 3:
        jpg_img = add_alpha_channel(jpg_img)
    
    '''
    当叠加图像时，可能因为叠加位置设置不当，导致png图像的边界超过背景jpg图像，而程序报错
    这里设定一系列叠加位置的限制，可以满足png图像超出jpg图像范围时，依然可以正常叠加
    '''
    yy1 = 0
    yy2 = png_img.shape[0]
    xx1 = 0
    xx2 = png_img.shape[1]

    if x1 < 0:
        xx1 = -x1
        x1 = 0
    if y1 < 0:
        yy1 = - y1
        y1 = 0
    if x2 > jpg_img.shape[1]:
        xx2 = png_img.shape[1] - (x2 - jpg_img.shape[1])
        x2 = jpg_img.shape[1]
    if y2 > jpg_img.shape[0]:
        yy2 = png_img.shape[0] - (y2 - jpg_img.shape[0])
        y2 = jpg_img.shape[0]

    # 获取要覆盖图像的alpha值，将像素值除以255，使值保持在0-1之间
    alpha_png = png_img[yy1:yy2,xx1:xx2,3] / 255.0
    alpha_jpg = 1 - alpha_png
    
    # 开始叠加
    for c in range(0,3):
        jpg_img[y1:y2, x1:x2, c] = ((alpha_jpg*jpg_img[y1:y2,x1:x2,c]) + (alpha_png*png_img[yy1:yy2,xx1:xx2,c]))

    return jpg_img

# 定义图像路径
img_jpg_path = bg_path # 读者可自行修改文件路径
ptt_png_path = ptt_pic_path # 读者可自行修改文件路径
head_path = os.path.join(current_file_dir + '/pictures', 'title.png')
f1_path = os.path.join(current_file_dir + '/pictures', 'f1.png')

# 读取图像
img_jpg = cv2.imread(img_jpg_path, cv2.IMREAD_UNCHANGED)
img_png = cv2.imread(ptt_png_path, cv2.IMREAD_UNCHANGED)
img_head = cv2.imread(head_path, cv2.IMREAD_UNCHANGED)
img_f1 = cv2.imread(f1_path, cv2.IMREAD_UNCHANGED)

# 设置叠加位置坐标
x1 = 1000
y1 = 100
x2 = x1 + img_png.shape[1]
y2 = y1 + img_png.shape[0]

# 叠加ptt框
res_img = merge_img(img_jpg, img_png, y1, y2, x1, x2)

# 叠加表头
res_img = merge_img(res_img, img_head, 0, img_head.shape[0], 120, 120+img_head.shape[1])
res_img = merge_img(res_img, img_f1, 230, 230+img_f1.shape[0], 400, 400+img_f1.shape[1])

bg2_path = os.path.join(current_file_dir + '/pictures', 'bg2.jpg')
# 保存结果图像
cv2.imwrite(bg2_path, res_img)

bg2 = Image.open(bg2_path)
# 新建一个空白图像
result = Image.new('RGBA', (img_width, img_height), (255, 255, 255))
result.paste(bg2, (0, 0))

# 右上文字（ptt）
ptt = str(round(float(info.iloc[0, 9]),2))
draw = ImageDraw.Draw(result)
draw.multiline_text((1025,150), ptt, font=ptt_font, fill=(255, 255, 255))
del draw

# 右上背景
drawRect(result,[(950, 250), (1190, 380)], fill=(10, 50, 200, 120))
# 右上文字（其他）
best30 = str(round(float(info.iloc[2, 9]),3))
pttmax = str(round(float(info.iloc[1, 9]),3))
r10 = str(round(float(info.iloc[3, 9]),3))
rinfo = "Best30: " + best30 + "\nRecent10: " + r10  + "\nPTTmax: " + pttmax
draw = ImageDraw.Draw(result)
draw.multiline_text((960,260), rinfo, font=ptt_font, fill=(255, 255, 255))
del draw

# 左上背景
drawRect(result,[(60, 220), (355, 380)], fill=(10, 50, 200, 120))

# 左上表头
now = datetime.datetime.now()
id = info.iloc[5, 9]
head = "Player: " + str(id) +"\nDate: " + str(now)[:10] + "\nTime: " + str(now)[11:16]
draw = ImageDraw.Draw(result)
draw.multiline_text((70,230), head, font=font, fill=(255, 255, 255))
del draw

# 右下标注

draw = ImageDraw.Draw(result)
draw.multiline_text((800, 1945), "Generated by AstronautDale", font=name_font, fill=(255,255,255))
del draw

# 排名
rank = ['01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30']


# 遍历曲目信息，将信息分为左右两列展示
for i in range(30):
    # 计算当前曲目所在方格的位置
    row = i // 4
    col = i % 4
    x0 = 60 + col * box_width
    y0 = 420 + row * box_height

    # 绘制方格背景
    drawRect(result,[(x0, y0), (x0 + 272, y0 + 170)], fill=(255, 255, 255, 128))

    # 获取曲目信息
    id = info.iloc[i, 1]
    name = info.iloc[i, 2]
    difficulty = info.iloc[i, 3]
    level = info.iloc[i, 4]
    score = info.iloc[i, 5]
    potential = str(round(float(info.iloc[i, 6]),3))
    
    # 加载曲目图片并将其缩放至方格大小
    song_img = Image.new('RGBA', (120, 120), (255, 255, 255))
    img_path = os.path.join(current_file_dir +'/songs', id + '.jpg')
    img = Image.open(img_path)
    img.thumbnail((120, 120))
    song_img.paste(img, (0, 0))

    font = ImageFont.truetype(font_path, size=28)
    # 在曲目图片上绘制排名
    drawRect(song_img,[(0, 0), (40, 30)], fill=(50, 50, 50, 128))
    draw = ImageDraw.Draw(song_img)
    draw.text((5, -5), str(rank[i]), font=font, fill=(255, 255, 255))
    del draw

    song_img_path = os.path.join(current_file_dir + '/pictures', 'song_img.png')
    song_img.save(song_img_path)

    song_img = cv2.imread(song_img_path, cv2.IMREAD_UNCHANGED)
    # 绘制难度
    if(str(difficulty) == 'FTR'):
        df_img_path = os.path.join(current_file_dir +'/pictures/tag-difficulty-future.png')
    elif(str(difficulty) == 'BYD'):
        df_img_path = os.path.join(current_file_dir +'/pictures/tag-difficulty-beyond.png')
    elif(str(difficulty) == 'PRS'):
        df_img_path = os.path.join(current_file_dir +'/pictures/tag-difficulty-present.png')
    elif(str(difficulty) == 'PST'):
        df_img_path = os.path.join(current_file_dir +'/pictures/tag-difficulty-past.png')
    df_img = cv2.imread(df_img_path, cv2.IMREAD_UNCHANGED)
    song_img = merge_img(song_img, df_img, 98, 98+df_img.shape[0], 47, 47+df_img.shape[1])
    cv2.imwrite(song_img_path, song_img)
    song_img = Image.open(song_img_path)

    # 将曲目图片放置到左侧方格
    result.paste(song_img, (x0+10, y0+10))

    font = ImageFont.truetype(font_path, size=32)
    # 将曲目得分、定数和潜力值放置到右侧方格
    text = str(score) + '\n' + str(level) + '\n' + str(potential)
    draw = ImageDraw.Draw(result)
    draw.multiline_text((x0 + 133, y0 + 14), text, font=score_font, fill=(0, 0, 0))
    del draw

    #将曲目名称绘制在下面的方格中
    draw = ImageDraw.Draw(result)
    draw.multiline_text((x0 + 5, y0 + 125), shorten_string(str(name)), font=font, fill=(0, 0, 0))
    del draw

# 保存结果图
result.save(current_file_dir +'/result.png')