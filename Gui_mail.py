"""
功能：选择文件，然后对文件进行处理并输出

最后编辑：2019年12月03日
"""

import sys
import os
import time
import EmailSend

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton,QFileDialog,QWidget,QLabel, QFrame,QProgressBar,QMessageBox
from PyQt5 import QtGui
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QRect,QCoreApplication,Qt,QBasicTimer
from PyQt5 import QtCore

#建立一个窗口类
class MyWindow(QMainWindow):
    """
    使用GUI界面选择附件、通讯录、邮件正文
    """
    def __init__(self):
        super().__init__()
        self.Init_Ui()

    #窗体UI部分
    def Init_Ui(self):
        #设置窗体大小
        self.setGeometry(500,300,400,300)
        #设置固定大小
        self.setFixedSize(500, 400)
        #设置标题
        self.setWindowTitle('提取文件进行处理')
        self.setWindowIcon(QIcon('icon/769160.png'))

        # 设置软件过期时间
        data = '2020-12-21 13:50:00'
        data_array = time.strptime(data, "%Y-%m-%d %H:%M:%S")
        timeStamp = int(time.mktime(data_array))

        # 判断软件是否过期
        if time.time() > timeStamp:
            QMessageBox.warning(self, '', '软件已过期，请联系作者', QMessageBox.Yes)
        else:
        ###===========设置标签===========================================
            #说明标签
            content = '说明:\n    请选择相应的附件、通讯录和图片。\n    如有问题，请联系IT同事！\n    联系电话：123456\n'
            self.description_lab = QLabel(content, self)
            self.description_lab.move(20, 10)
            self.description_lab.setFixedSize(460, 80)
            self.description_lab.setAlignment(Qt.AlignTop)
            self.description_lab.setFrameShape(QFrame.Box)
            self.description_lab.setFrameShadow(QFrame.Raised)
            # 自动换行
            # self.description_lab.adjustSize()
            self.description_lab.setWordWrap(True)

            # 正文图片路径标签
            self.lab_img_path = QtWidgets.QLabel(self)
            self.lab_img_path.move(150, 120)
            self.lab_img_path.setFixedSize(330,40)
            self.lab_img_path.setFrameShape(QFrame.Box)
            self.lab_img_path.setFrameShadow(QFrame.Raised)              
            # 邮件正文路径标签
            self.lab_mailtext_path = QtWidgets.QLabel(self)
            self.lab_mailtext_path.move(150, 170)
            self.lab_mailtext_path.setFixedSize(330,40)
            self.lab_mailtext_path.setFrameShape(QFrame.Box)
            self.lab_mailtext_path.setFrameShadow(QFrame.Raised)  
            # 附件路径标签
            self.lab_attach_path = QtWidgets.QLabel(self)
            self.lab_attach_path.move(150, 220)
            self.lab_attach_path.setFixedSize(330,40)
            self.lab_attach_path.setFrameShape(QFrame.Box)
            self.lab_attach_path.setFrameShadow(QFrame.Raised)
            # 通讯录路径标签
            self.lab_addrfile_path = QtWidgets.QLabel(self)
            self.lab_addrfile_path.move(150, 270)
            self.lab_addrfile_path.setFixedSize(330,40)
            self.lab_addrfile_path.setFrameShape(QFrame.Box)
            self.lab_addrfile_path.setFrameShadow(QFrame.Raised)   

        ###===========设置按钮===========================================
            #选择图片按钮
            self.btn_img_file = QPushButton('选择图片',self)
            self.btn_img_file.move(20,120)
            self.btn_img_file.setFixedSize(120,40)
            # self.btn_selectattach_file.setGeometry(QRect(50,100,120,40))
            self.btn_img_file.clicked.connect(self.SelectImg)        
            #选择邮件正文按钮
            self.btn_mailtext_file = QPushButton('选择Html',self)
            self.btn_mailtext_file.move(20,170)
            self.btn_mailtext_file.setFixedSize(120,40)
            # self.btn_selectattach_file.setGeometry(QRect(50,100,120,40))
            self.btn_mailtext_file.clicked.connect(self.SelectMailText)
            #选择附件按钮
            self.btn_selectattach_file = QPushButton('选择附件',self)
            self.btn_selectattach_file.move(20,220)
            self.btn_selectattach_file.setFixedSize(120,40)
            # self.btn_selectattach_file.setGeometry(QRect(50,100,120,40))
            self.btn_selectattach_file.clicked.connect(self.SelectAttachfile)
            #选择通讯录按钮
            self.btn_selectaddr_file = QPushButton('选择通讯录',self)
            self.btn_selectaddr_file.move(20,270)
            self.btn_selectaddr_file.setFixedSize(120,40)
            self.btn_selectaddr_file.clicked.connect(self.SelectAddr)

            #开始按钮
            self.start_button = QPushButton('开始', self)
            self.start_button.setGeometry(QRect(20,320,220,40))
            # 点击按钮，连接事件函数
            self.start_button.clicked.connect(self.btn_action)

            #结束按钮
            self.exit_button = QPushButton('结束', self)
            self.exit_button.setGeometry(QRect(260,320,220,40))
            # self.exit_button.resize(self.exit_button.sizeHint())
            # btn.move(50, 50)
            self.exit_button.clicked.connect(QCoreApplication.instance().quit)


    ###===========每个按钮对应的函数==============================================
    #选择图片
    def SelectImg(self):
        img_name = QFileDialog.getOpenFileName(self,"选择邮件正文",'',"Picture Files (*.jpeg)")
        self.lab_img_path.setText(img_name[0])
        global image_name
        image_name = img_name[0]

    #选择邮件正文
    def SelectMailText(self):
        #返回的是一个元组，[0]提取，或者分别赋值都可以
        text_name = QFileDialog.getOpenFileName(self,"选择邮件正文",'',"Html Files (*.html)")
        # 标签框显示文本路径
        self.lab_mailtext_path.setText(text_name[0])
        # 自动调整标签框大小
        #self.lab_attach_path.adjustSize()
        #定义一个全局变量
        global mailtext_name
        #获取文件名称
        mailtext_name = text_name[0]

    #选择附件路径
    def SelectAttachfile(self):
        path = QFileDialog.getExistingDirectory(self,"选择附件",'')
        self.lab_attach_path.setText(path)
        global attachfile_name
        attachfile_name = path

    # 选择通讯录文件
    def SelectAddr(self):
        addr_file = QFileDialog.getOpenFileName(self, '选取文件', '',"CSV file(*.csv)")
        # 标签框显示文本路径
        self.lab_addrfile_path.setText(addr_file[0])
        global addrfile_name
        addrfile_name = addr_file[0]


    # 按钮点击
    def btn_action(self,type):
        if self.start_button.text() == '完成':
            self.close()
        else:
            html_path = '{}'.format(self.lab_mailtext_path.text())
            attachfile_path = '{}'.format(self.lab_attach_path.text())
            addrfile_path = '{}'.format(self.lab_addrfile_path.text())
            img_path = '{}'.format(self.lab_img_path.text())
            if self.start_button.text() == '开始':
                    if not os.path.exists(html_path):
                        QMessageBox.warning(self, '', '请选择邮件正文', QMessageBox.Yes)
                    elif not os.path.exists(attachfile_path):
                        QMessageBox.warning(self, '', '请选择附件路径', QMessageBox.Yes)
                    elif not os.path.exists(addrfile_path):
                        QMessageBox.warning(self, '', '请选择通讯录', QMessageBox.Yes)
                    elif not os.path.exists(img_path):
                        QMessageBox.warning(self, '', '请选择图片', QMessageBox.Yes)
                    else:
                        self.start_button.setText('程序进行中')
                        EmailSend.send(attachfile_path,addrfile_path,html_path,img_path)
                        self.start_button.setText('完成')
    

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show() 
    sys.exit(app.exec_())
    