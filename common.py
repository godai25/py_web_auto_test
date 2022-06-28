import os
import datetime

class common_method :
    # ログ出力
    def write_log(msg:str) -> None:
        with open(os.path.basename(__file__) + ".log", mode="a", encoding="Shift-JIS") as f:
            print(msg)
            now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            f.writelines(now + " " + msg + "\n")


