import os

# 定数扱いの全部をここで宣言
class all_const :
    CHROME_DRIVER :str = "C:\work\python_test\chromedriver.exe"

    # 「main.py」があるフォルダの場所
    BASE_DIR = os.getcwd() + "\\"

    EXCEL_DIR = BASE_DIR + "Excel\\"
    IMAGE_DIR = BASE_DIR + "Img\\"
    TEMP_DIR =  BASE_DIR + "Temp\\"
