import sys
sys.path.append("T:/development/programgarden/pro_garden")

from pro_garden.kiwoom.kiwoom import *
from PyQt5.QtWidgets import *

class Main():
    def __init__(self):
        print("Main() start")

        self.app = QApplication(sys.argv) # PyQt5로 실행할 파일명을 자동 설정
        self.kiwoom = Kiwoom() # 키움 클래스 실행
        self.app.exec_() # 이벤트 루프 실행


if __name__ == "__main__":
    Main()
