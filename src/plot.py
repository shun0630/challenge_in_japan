import sys
import openpyxl
import os
import shutil
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import textwrap
import json


#エクセルからデータを読み込む
def input_data(file_name):
    wb = openpyxl.load_workbook(file_name, data_only=True)
    sheetnames_list = wb.sheetnames
    for sheetname in sheetnames_list:
        sheet = wb.get_sheet_by_name(str(sheetname))
        rank_data = []
        group_data = []

        for row in range(5,sheet.max_row+1):
            event = sheet['a' + str(3)].value
            grade = sheet['b' + str(3)].value
            rank = sheet['a' + str(row)].value
            name = sheet['b' + str(row)].value
            group = sheet['c' + str(row)].value
            count = sheet['d' + str(row)].value


            data = {"event":event,"grade":grade ,"rank":"第" + str(rank) + "位", "name":name, "group":group, "count":"記録：" + str(count) + "回"}
            rank_data.append(data)
            group_data.append(group)
    group_data_list = set(group_data)
    team_data_list = {"rank_data":rank_data, "group_data":group_data_list}
    return team_data_list
    

#団体ごとのフォルダを作成   
def mkdir_group(group_data_list):
    shutil.rmtree('./sankasho')
    os.mkdir("./sankasho")
    for group in group_data_list:
        os.mkdir('./sankasho/' + str(group))




#フォントの設定
def font_setting(font_data, team_data, im):
    font_data_list = {}
    for data in font_data:
        font_data_list[data] =  font_position(team_data[data], font_data[data]["size"], font_data[data]["height"], font_data[data]["width"], im)
    return font_data_list



#書き込み場所、フォントサイズの設定
def font_position(write_comment, font_size, font_height_input, font_width_input, im):
    im_copy = im
    draw = ImageDraw.Draw(im_copy)
    font_path = "/usr/share/fonts/truetype/fonts-japanese-mincho.ttf"
    W_im, H_im = im.size
    font = ImageFont.truetype(font_path, int(font_size))
    w_text, h_text = draw.textsize(str(write_comment), font)
    print(w_text, h_text)
    
    #文字数に合わせてフォントサイズを調整する。
    while W_im - w_text < W_im/1.7:
        font_size = font_size - 5
        font = ImageFont.truetype(font_path, int(font_size))
        w_text, h_text = draw.textsize(str(write_comment), font)
    font_height = font_height_input
    font_width = W_im/2
    return {"write_comment": write_comment,"font": font, "font_height":font_height, "font_width": font_width}


#画像にテキストを書き込み
def image_write(data, im, font_data_list):
    im_copy = im
    draw = ImageDraw.Draw(im_copy)
    for write_item in font_data_list:
        draw.text((font_data_list[write_item]["font_width"], font_data_list[write_item]["font_height"]), str(data[write_item]), fill = "black", font = font_data_list[write_item]["font"], anchor='mm')
    im_copy.save("./sankasho/" + str(data["group"]) + "/" + str(data["name"]) + ".png")    



def main():

    exel_file_name =  sys.argv[1]
    image_file = sys.argv[2]
    font_data_json = sys.argv[3]

    #フォントパラメータの読み込み（サイズ等）
    with open(font_data_json) as f:
        font_data = json.load(f)

    # データをエクセルから読み込み
    team_data_list = input_data(exel_file_name)

    #団体ごとのフォルダ作成
    mkdir_group(team_data_list["group_data"])

    #書き込み用画像の用意
    im = Image.open(image_file)

    for team_data in team_data_list["rank_data"]:
        #書き込み用データのセット
        im = Image.open(image_file)

        #フォントの設定
        font_data_list = font_setting(font_data, team_data, im)

        # 参加賞画像に書き込み
        image_write(team_data, im, font_data_list)
if __name__ == "__main__":
    main()