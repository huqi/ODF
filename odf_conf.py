from docx.enum.text import WD_ALIGN_PARAGRAPH


STR_WINDOW_TITLE = "公文格式化"
STR_WARNING = "警告"
STR_ERROR = "错误"
STR_DONE = "完成"
STR_ALREADY_DONE = "已完成，格式化后的文件名为{0}"
STR_UNKNOWN = "未知错误"
STR_LOAD_CONF_ERR = "未找到配置文件，已将配置文件恢复为默认值"
STR_RUNNING_ERR = "有任务正在运行"
STR_DOC_TYPE_ERR = "文件类型错误，只支持doc/docx文件类型"
STR_LABEL_DOCX = "正在转换为DOCX文件……"
STR_LABEL_REPLACE_SPACE = "正在处理空格……"
STR_LABEL_REPLACE_BLANK = "正在处理空行……"
STR_LABEL_FORMAT = "正在处理格式……"
STR_LABEL_WRITE = "正在写入……"

MAIN_FRAME_SIZE = {'WIDTH': 300, 'HEIGHT': 400}
CONF_FILENAME = "conf.json"

IMG_DRAG_DROP_BLUE_PATH = './src/drag-drop-blue.png'
IMG_DRAG_DROP_GRAY_PATH = './src/drag-drop-grey.png'
IMG_ICON_PATH = './src/icon.png'

STEP = ['replace_space', 'replace_blank', 'format', 'write']

P_STYLE = {
    'b_title': {
        'font': '方正楷体_GBK',
        'font_size': '三号',
        'indent': 0,
        'align': '两端',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8,
        'line_count': 2
    },
    'title': {
        'font': '方正小标宋_GBK',
        'font_size': '二号',
        'indent': 0,
        'align': '居中',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'a_title': {
        'font': '方正楷体_GBK',
        'font_size': '三号',
        'indent': 0,
        'align': '居中',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8,
        'line_count': 2
    },
    'target': {
        'font': '方正仿宋_GBK',
        'font_size': '三号',
        'indent': 0,
        'align': '居左',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'head1': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent': 2,
        'align': '两端',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'head2': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent': 2,
        'align': '两端',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'head3': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent': 2,
        'align': '两端',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'head4': {
        'font': '方正黑体_GBK',
        'font_size': '三号',
        'indent': 2,
        'align': '两端',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'tail': {
        'font': '方正仿宋_GBK',
        'font_size': '三号',
        'indent': 0,
        'align': '居右',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8,
        'line_count': 2
    },
    'text': {
        'font': '方正仿宋_GBK',
        'font_size': '三号',
        'indent': 2,
        'align': '两端',
        'space_before': 0,
        'space_after': 0,
        'line_spacing': 28.8
    },
    'default': {
        'b_title_checked': False,
        'a_title_checked': False,
        'target_checked': True,
        'tail_checked': True
    }
}

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

FONT_SIZE_PT = {'二号': 22, '三号': 16, '四号': 14, '小四': 12, '五号': 10.5}
PARA_ALIGN = {'居左': WD_ALIGN_PARAGRAPH.LEFT,
              '居中': WD_ALIGN_PARAGRAPH.CENTER,
              '居右': WD_ALIGN_PARAGRAPH.RIGHT,
              '两端': WD_ALIGN_PARAGRAPH.JUSTIFY}
