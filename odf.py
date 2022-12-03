import codecs
import json
import os
import sys
import time

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QGuiApplication, QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QWidget
from win32com import client as wc

from odf_conf import *
from odf_ui import Ui_Form


def get_file_extension(file):
    return os.path.splitext(file)[-1]


def is_word_file(file):
    return any(file.endswith(ext) for ext in ['.doc', '.docx', 'DOC', '.DOCX'])


def doc_to_docx(file):
    word = wc.Dispatch("Word.Application")
    doc = word.Documents.Open(file)
    doc.SaveAs("{}x".format(new_file(file)), 12)
    doc.Close()  # 关闭原来word文件
    word.Quit()


def get_time_str():
    return time.strftime("%Y%m%d%H%M%S", time.localtime())


def new_file(oldfile):
    path, file = os.path.split(oldfile)
    filename, ext = os.path.splitext(file)
    new_filename = get_time_str() + "-" + filename + ext

    return os.path.join(path, new_filename)


class MyMainForm(QWidget, Ui_Form):
    def __init__(self, parent=None):
        '''
        构造函数
        '''
        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)

        self.img_blue = QPixmap(IMG_DRAG_DROP_BLUE_PATH)
        self.img_grey = QPixmap(IMG_DRAG_DROP_GRAY_PATH)

        self.setWindowTitle(STR_WINDOW_TITLE)
        self.setWindowIcon(QIcon(IMG_ICON_PATH))

        x, y = self.getCenterPos()

        self.setGeometry(
            x, y, MAIN_FRAME_SIZE['WIDTH'], MAIN_FRAME_SIZE['HEIGHT'])
        self.setFixedSize(self.width(), self.height())
        self.label_button.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)
        self.eventRestore(None, None, None)

    def setConfig(self, conf):
        self.conf = conf

    def getCenterPos(self):
        '''
        获得居中坐标
        '''
        screen = QGuiApplication.primaryScreen().size()

        x = int((screen.width() - MAIN_FRAME_SIZE['WIDTH']) / 2)
        y = int((screen.height() - MAIN_FRAME_SIZE['HEIGHT']) / 2)

        return x, y

    def dragEnterEvent(self, event):
        '''
        拖入
        '''
        event.accept()
        self.label_button.setPixmap(self.img_blue)

    def dragMoveEvent(self, event):
        '''
        移动
        '''
        pass

    def dragLeaveEvent(self, event):
        '''
        移出
        '''
        self.label_button.setPixmap(self.img_grey)

    def showMessageBox(self, type, title, msg):
        if type == "warning":
            QMessageBox.warning(self, title, msg)


    def eventRestore(self, type, title, message):
        '''
        进度条重置
        '''
        self.showMessageBox(type, title, message)
        self.label_button.setPixmap(self.img_grey)
        self.progressBar.setValue(0)
        self.label_progress.setText("(0/0)")
        self.running = False

    def evenErrorCallback(self):
        self.eventRestore("warning", STR_ERROR, STR_UNKNOWN)
        
    def evenLabelCallback(self, i, count):
        self.label_progress.setText("({0}/{1})".format(i, count))
        
    def getAbility(self):
        
        ability = {'checkBox_b_title' : False, 
                   'checkBox_a_title' : False, 
                   'checkBox_target' : False, 
                   'checkBox_tail' : False}
        
        if self.checkBox_b_title.isChecked():
            ability['checkBox_b_title'] = True
            
        if self.checkBox_b_title.isChecked():
            ability['checkBox_a_title'] = True
            
        if self.checkBox_target.isChecked():
            ability['checkBox_target'] = True
            
        if self.checkBox_target.isChecked():
            ability['checkBox_tail'] = True
            
        return ability
            

    def dropEvent(self, event):
        '''
        松开鼠标
        '''
        if self.running:
            self.eventRestore("warning", STR_ERROR, STR_RUNNING_ERR)
            return

        file = event.mimeData().urls()[0].toLocalFile()

        if not is_word_file(file):
            self.eventRestore("warning", STR_ERROR, STR_DOC_TYPE_ERR)
            return

        self.running = True
        self.odf = ODFThread(file, self.conf, self.getAbility())
        self.odf.sig_err.connect(self.evenErrorCallback)
        self.odf.sig_label.connect(self.evenLabelCallback)
        self.odf.run()
        # odf.start()


class ODFThread(QThread):
    '''
    image to pdf 线程
    '''
    sig_err = pyqtSignal()
    sig_label = pyqtSignal(int, int)
 
    def __init__(self, file, conf, ability):
        '''
        构造函数
        '''

        super(ODFThread, self).__init__()
        self.file = file
        self.conf = conf
        self.ability = ability

    def getStepCount(self):
        c = 0
        
        for a, s in self.ability.items():
            if a == 'checkBox_b_title' and s:
                c += 2
                
            if a == 'checkBox_a_title' and s:
                c += 2

        return 0

    def run(self):
        '''
        线程启动函数
        '''
        self.getStepCount()

        try:
            if get_file_extension(self.file) in [".doc", ".DOC"]:
                doc_to_docx(self.file)

        except:
            self.sig_err.emit()

def load_conf(filename):
    with codecs.open(filename, 'r', "utf-8") as f:
        return json.load(f)


def restore_conf(filename, data):
    with codecs.open(filename, "w", "utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    return data


def get_conf(self):

    data = None

    try:
        data = load_conf(CONF_FILENAME)
    except:
        QMessageBox.warning(self, STR_WARNING, STR_LOAD_CONF_ERR)
        data = restore_conf(CONF_FILENAME, P_STYLE)

    return data


def main():
    # '''
    # 主函数
    # '''
    app = QApplication(sys.argv)
    myWin = MyMainForm()
    myWin.show()
    myWin.setConfig(get_conf(myWin))

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()


b_title_flag = True
b_title_line_count = 2
a_title_flag = True
a_title_line_count = 2
target_line_flag = True
tail_line_flag = True
tail_line_count = 2


class Para:
    text = ""
    content = ""
    type = ""

    def __init__(self, text, type=None):
        if not type:
            self.type = self.get_type(text)
        else:
            self.type = type

        if self.type in ('head2', 'head3', 'head4'):
            self.text, self.content = text.split("。", 1)
            self.text += "。"
        else:
            self.text = text

    def get_type(self, text):
        for k, v in HEAD_SIGN.items():
            for sign in v:
                if text.startswith(sign):
                    return k

        return 'text'

    def print_type(self):
        print(self.type)

    def print_text(self):
        print(self.text)


class Doc:
    paragraphs = []
    document = None
    doc = []
    filename = ""
    line_no = 0
    b_title_done = False

    def __init__(self, file):
        self.document = Document(file)
        self.filename = file

    def print(self):
        for p in self.paragraphs:
            p.print_type()
            p.print_text()

    def wash_line(self):
        para = []
        for line in self.doc:
            if len(line) > 0:
                para.append(line)
                # self.paragraphs.append(Para(line))

        return para

    def load_para(self, para):
        b_title_c = 0
        title_done = False
        a_title_c = 0
        target_line_done = False
        tail_line_c = 0
        tail_blank_line_done = False
        line_no = 0

        for line in para:
            line_no += 1

            if b_title_flag and b_title_c < b_title_line_count:
                self.paragraphs.append(Para(line, "b_title"))
                b_title_c += 1

                if b_title_c == b_title_line_count:
                    self.paragraphs.append(Para("", "text"))

                continue

            if not title_done:
                self.paragraphs.append(Para(line, "title"))
                if not a_title_flag:
                    self.paragraphs.append(Para("", "text"))

                title_done = True
                continue

            if a_title_flag and a_title_c < a_title_line_count:
                self.paragraphs.append(Para(line, "a_title"))
                a_title_c += 1

                if a_title_c == a_title_line_count:
                    self.paragraphs.append(Para("", "text"))

                continue

            if target_line_flag and not target_line_done:
                self.paragraphs.append(Para(line, "target"))
                target_line_done = True
                continue

            if tail_line_flag and line_no > len(para) - tail_line_count and tail_line_c < tail_line_count:
                if not tail_blank_line_done:
                    self.paragraphs.append(Para("", "text"))
                    self.paragraphs.append(Para("", "text"))
                    tail_blank_line_done = True

                self.paragraphs.append(Para(line, "tail"))
                tail_line_c += 1
                continue

            self.paragraphs.append(Para(line))

    def replace_x(self, text):
        t = text
        for x in x_list:
            t = t.replace(x, '')

        return t

    def replace_blank(self):
        for line in self.document.paragraphs:
            self.doc.append(self.replace_x(line.text))

    def run(self):
        self.replace_blank()
        self.load_para(self.wash_line())

    def getTimeStr(self):
        return time.strftime("%Y%m%d%H%M%S", time.localtime())

    def write(self):
        doc = Document()
        for line in self.paragraphs:

            p = doc.add_paragraph()
            run = p.add_run(line.text)
            run.font.name = 'Times New Roman'
            run.element.rPr.rFonts.set(
                qn('w:eastAsia'), P_STYLE[line.type]['font'])
            run.font.size = Pt(FONT_SIZE_PT[P_STYLE[line.type]['font_size']])

            if line.content:
                run_content = p.add_run(line.content)
                run_content.font.name = 'Times New Roman'
                run_content.element.rPr.rFonts.set(
                    qn('w:eastAsia'), P_STYLE['text']['font'])
                run_content.font.size = Pt(
                    FONT_SIZE_PT[P_STYLE['text']['font_size']])

            p.paragraph_format.space_before = Pt(
                P_STYLE[line.type]['space_before'])
            p.paragraph_format.space_after = Pt(
                P_STYLE[line.type]['space_after'])
            p.paragraph_format.line_spacing = Pt(
                P_STYLE[line.type]['line_spacing'])
            p.paragraph_format.alignment = PARA_ALIGN[P_STYLE[line.type]['align']]
            p.paragraph_format.first_line_indent = run.font.size * \
                P_STYLE[line.type]['indent']

        doc.save("t.docx")


# doc = Doc("test.docx")
# doc.run()
# doc.print()
# doc.write()
