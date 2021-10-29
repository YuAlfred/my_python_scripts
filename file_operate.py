import os
import shutil

# 遍历文件夹并根据文件名对文件进行分类
print(os.getcwd())

rootDir = "./中小学/"

# 创建w
category = ["小学", "初级中学", "九年一贯制学校", "十二年一贯制学校", "完全中学"]
for c in category:
    if not os.path.exists(rootDir + c):
        os.makedirs(rootDir + c)

primary_school = "小学"
middle_school = "初级中学"
high_school = "高中"
nine_year_school = "九年一贯制学校"
twelve_year_school = "十二年一贯制学校"
total_school = "完全中学"

i = 1
for f in os.listdir(rootDir):
    if os.path.isfile(rootDir + f):
        print("正在处理：" + f)
        for c in category:
            if c in f:
                shutil.copy(rootDir + f, rootDir + c + "/" + f)
                print("处理完成：" + f + "  编号:" + str(i))
                i = i + 1
                break
