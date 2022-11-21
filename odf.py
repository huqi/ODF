from docx import Document

x_dict = {',' : '，' ,
          '.' : '。' ,
          '!' : '！' ,
          '(' : '（' ,
          ')' : '）' }

x_list = ['\t', ' ']

head1 = ['一、', '二、', '三、', '四、', '五、',
         '六、', '七、', '八、', '九、', '十、',
         '十一、', '十二、', '十三、', '十四、', '十五、',
         '十六、', '十七、', '十八、', '十九、', '二十、']

head2 = ['（一）', '（二）', '（三）', '（四）', '（五）',
         '（六）', '（七）', '（八）', '（九）', '（十）',
         '（十一）', '（十二）', '（十三）', '（十四）', '（十五）',
         '（十六）', '（十七）', '（十八）', '（十九）', '（二十）']

head3 = ['1.' , '2.' , '3.' , '4.' , '5.' ,
         '6.' , '7.' , '8.' , '9.' , '10.',
         '11.', '12.', '13.', '14.', '15.',
         '16.', '17.', '18.', '19.', '20.']

head4 = ['（1）' , '（2）' , '（3）' , '（4）' , '（5）' ,
         '（6）' , '（7）' , '（8）' , '（9）' , '（10）',
         '（11）', '（12）', '（13）', '（14）', '（15）',
         '（16）', '（17）', '（18）', '（19）', '（20）']

head = {'head1' : head1, 
        'head2' : head2,
        'head3' : head3,
        'head4' : head4}

def is_head(text):
    for k, v in head.items():
        for sign in v:
            if text.startswith(sign):
                return k
            
    return 'text'

def wash_x(text):
    
    t = text
    for x in x_list:
        t = t.replace(x, '')
        
    return t

def wash_line(page):
    
    p = []
    
    for line in page:
        if len(line) > 0:
            p.append(line)
            
    return p
            
def del_blank(page):
    
    p = []
    
    for line in page:
        p.append(wash_x(line.text))
        
    return p

def print_page(page):
    for line in page:
        #print(page)
        print(is_head(line))
        

document = Document("test.docx")
para = document.paragraphs
para = del_blank(para)
para = wash_line(para)
print_page(para)