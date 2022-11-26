import time

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


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

    def __init__(self, text, type = None):
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
                #self.paragraphs.append(Para(line))
        
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
            
            if b_title_flag and  b_title_c < b_title_line_count:
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
                
            if tail_line_flag and line_no > len(para) - tail_line_count and tail_line_c < tail_line_count :
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
            run.element.rPr.rFonts.set(qn('w:eastAsia'), P_STYLE[line.type]['font'])
            run.font.size = Pt(FONT_SIZE_PT[P_STYLE[line.type]['font_size']])
            
            if line.content:
                run_content = p.add_run(line.content)
                run_content.font.name = 'Times New Roman'
                run_content.element.rPr.rFonts.set(qn('w:eastAsia'), P_STYLE['text']['font'])
                run_content.font.size = Pt(FONT_SIZE_PT[P_STYLE['text']['font_size']])

            p.paragraph_format.space_before = Pt(P_STYLE[line.type]['space_before'])
            p.paragraph_format.space_after = Pt(P_STYLE[line.type]['space_after'])
            p.paragraph_format.line_spacing = Pt(P_STYLE[line.type]['line_spacing'])
            p.paragraph_format.alignment = PARA_ALIGN[P_STYLE[line.type]['align']]
            p.paragraph_format.first_line_indent = run.font.size * P_STYLE[line.type]['indent']

        doc.save("t.docx")


x_dict = {
    ',': '，',
    '.': '。',
    '!': '！',
    '(': '（',
    ')': '）'
}

x_list = ['\t', ' ']

head1 = ['一、', '二、', '三、', '四、', '五、',
         '六、', '七、', '八、', '九、', '十、',
         '十一、', '十二、', '十三、', '十四、', '十五、',
         '十六、', '十七、', '十八、', '十九、', '二十、']

head2 = ['（一）', '（二）', '（三）', '（四）', '（五）',
         '（六）', '（七）', '（八）', '（九）', '（十）',
         '（十一）', '（十二）', '（十三）', '（十四）', '（十五）',
         '（十六）', '（十七）', '（十八）', '（十九）', '（二十）']

head3 = ['1.', '2.', '3.', '4.', '5.',
         '6.', '7.', '8.', '9.', '10.',
         '11.', '12.', '13.', '14.', '15.',
         '16.', '17.', '18.', '19.', '20.']

head4 = ['（1）', '（2）', '（3）', '（4）', '（5）',
         '（6）', '（7）', '（8）', '（9）', '（10）',
         '（11）', '（12）', '（13）', '（14）', '（15）',
         '（16）', '（17）', '（18）', '（19）', '（20）']

HEAD_SIGN = {'head1': head1,
             'head2': head2,
             'head3': head3,
             'head4': head4}

FONT_SIZE_PT = {'二号': 22, '三号' : 16, '四号' : 14, '小四' : 12, '五号' : 10.5}
PARA_ALIGN = {'居左' : WD_ALIGN_PARAGRAPH.LEFT, 
              '居中' : WD_ALIGN_PARAGRAPH.CENTER,
              '居右' : WD_ALIGN_PARAGRAPH.RIGHT,
              '两端' : WD_ALIGN_PARAGRAPH.JUSTIFY}


P_STYLE = {
    'b_title': {
        'font': '方正楷体_GBK',
        'font_size': '三号',
        'indent' : 0,
        'align' : '两端',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'title': {
        'font': '方正小标宋_GBK',
        'font_size': '二号',
        'indent' : 0,
        'align' : '居中',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'a_title': {
        'font': '方正楷体_GBK',
        'font_size': '三号',
        'indent' : 0,
        'align' : '居中',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'target': {
        'font': '方正仿宋_GBK',
        'font_size': '三号',
        'indent' : 0,
        'align' : '居左',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },    
    'head1': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent' : 2,
        'align' : '两端',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'head2': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent' : 2,
        'align' : '两端',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'head3': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent' : 2,
        'align' : '两端',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'head4': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent' : 2,
        'align' : '两端',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'tail': {
        'font': '方正仿宋_GBK',
        'font_size': '三号',
        'indent' : 0,
        'align' : '居右',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    },
    'text': {
        'font': '方正仿宋_GBK',
        'font_size': '三号',
        'indent' : 2,
        'align' : '两端',
        'space_before' : 0,
        'space_after' : 0,
        'line_spacing' : 28.8
    }
}


doc = Doc("test.docx")
doc.run()
doc.print()
doc.write()
