import streamlit as st
import logging
import mechanize
from bs4 import BeautifulSoup
import urllib.request
import pandas
import requests
import sys, time, msvcrt
import re
from os.path import exists
from unidecode import unidecode
import openpyxl


f = open("Log.txt", "a", encoding="utf-8")

class Books:
    def __init__(self):
        self.Name = ''           #عنوان
        self.Author = ''         #نویسنده
        self.Publisher = ''      #ناشر
        self.Series = ''         #فروست
        self.Illustrator = ''    #تصویرگر
        self.Translator = ''     #مترجم
        self.ISBN = ''           #شابک
        self.Topics = ''         #موضوع
        self.Editor = ''         #ویراستار
        self.Year = ''           #سال اولین نشر
        self.City = ''           #شهر نشر
        self.Notes = ''          #یادداشت
        self.Dewey = ''          #رده بندی دیویی
        self.Congress = ''       #رده بندی کنگری
        self.NationalNo = ''     #کتابشناسی ملی
        self.AddedInfo = ''      #شناسه افزوده
    def __str__(self):
        return self.Name + '/' + self.Author + '/' + self.Publisher + '/' + self.ISBN
def readInput( caption, default, timeout = 20):
    start_time = time.time()
    print(caption, "(", default, ")")
    print(caption, "(", default, ")", file=f,  flush=True)
    input = ''
    while True:
        if msvcrt.kbhit():
            chr = msvcrt.getche()
            if ord(chr) == 13: # enter_key
                break
            elif ord(chr) == 8:
                if len(input) > 0:
                    input = input[0:-1]
            elif ord(chr) >= 32: #space_char
                try:
                    input += chr.decode("utf-8") 
                except:
                    pass
        if len(input) == 0 and (time.time() - start_time) > timeout:
            break

    print('')  # needed to move to next line
    if len(input) > 0:
        return input
    else:
        return default


def select_form(form):
    return form.attrs.get('action', None) == '/advanced_search'


def SearchInNLAI(*, ISBN=None, ISBNType="selected", \
                 Title=None, TitleType="selected", \
                 Author=None, AuthorType="selected", \
                 Publisher=None, PublisherType="selected", \
                 YearStart=None, YearEnd=None, \
                 Topic=None, TopicType="selected"):
    br = mechanize.Browser()
    br.set_handle_robots(False)
    br.addheaders.append(
        ('User-Agent',
         'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/601.3.9 (KHTML, like Gecko) Version/9.0.2 Safari/601.3.9'))
    headers = headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Cafari/537.36'}

    Check = False
    for i in range(0,3):
        try:
            br.open("https://libs.nlai.ir/advanced")
            Check = True
            break
        except:
            print("Failed To Open; Retrying ...", file=f,  flush=True)
    if ( not Check):
        return -2
    response = br.response()

    br.select_form(predicate=select_form)
    InputInfo = ''
    if Title != None:
        br.form.set_value(Title, id='searchTitle')
        InputInfo = InputInfo + ' - Title = '+Title
    if Author != None:
        br.form.set_value(Author, id='searchAuthor')
        InputInfo = InputInfo + ' - Author= '+Author

    if Publisher != None:
        br.form.set_value(Publisher, id='searchPublisher')
        InputInfo = InputInfo + ' - Publisher= '+Publisher
    if YearStart != None:
        InputInfo = InputInfo + ' - YeartStart= '+YearStart
        try:
            Year = int(YearStart)
            if Year > 1000:
                br.form.set_value(Year.str(), id='searchPublishYear')
        except ValueError:
            if YearStart.isnumeric():
                if int(YearStart) > 1000:
                    br.form.set_value(YearStart, id='searchPublishYear')

    if YearEnd != None:
        InputInfo = InputInfo + ' - YearEnd = '+YearEnd
        try:
            Year = int(YearEnd)
            if Year > 1000:
                br.form.set_value(Year.str(), name='publish_year_end')
        except ValueError:
            if YearEnd.isnumeric():
                if int(YearEnd) > 1000:
                    br.form.set_value(YearEnd, id='publish_year_end')

    if ISBN != None:
        InputInfo = InputInfo + ' - ISBN = '+ISBN
        br.form.set_value(ISBN, id='searchISBN')

    if Topic != None:
        InputInfo = InputInfo + ' - Topic = '+Topic
        br.form.set_value(Topic, id='searchTopic')
    st.text("پس دنبال این می‌گردم: " + InputInfo)
    try:
        br.submit()
    except:
        print("Failed To Submit" + InputInfo, file=f,  flush=True)
        return -2

    response2 = br.response()

    soup = BeautifulSoup(response2, "html.parser")

    checkboxes = soup.select(".form-check-input")
    TempFile = []
    TextToShow = []
    st.text(len(checkboxes))
    if len(checkboxes) == 0:
        print("No book found with input information"+InputInfo, file=f,  flush=True)
        return -1
    else:
        for idx, ck in enumerate(checkboxes):
            FileName = "RawExcels/"+ ck.get("value") + ".xlsx"
            if not(exists(FileName)):
                Test = False
                for i in range(0,3):
                    # urllib.request.urlretrieve("https://libs.nlai.ir/result/excel?uid="+ck.get("value"), [ck.get("value") + ".xlsx"])
                    try:
                        resp = requests.get("https://libs.nlai.ir/result/excel?uid=" + ck.get("value"), headers=headers)
                        Test = True
                    except:
                        pass
                if (not Test):
                    return -2
                output = open(FileName, 'wb')
                output.write(resp.content)
                output.close()
            TempFile.append(pandas.read_excel(FileName, index_col=None))
        return TempFile

def AnalyzeExcelRaw(FileExcelRaws):
    ResultingBooks = []
    for FileExcelRaw in FileExcelRaws:
        Book = Books()
        if "عنوان و نام پدیدآور" in FileExcelRaw:
            FileExcelRaw["عنوان و نام پدیدآور"][0] = FileExcelRaw["عنوان و نام پدیدآور"][0].replace('ي','ی')
            slashloc = FileExcelRaw["عنوان و نام پدیدآور"][0].rfind('/')
            if slashloc > 0:
                Book.Name = FileExcelRaw["عنوان و نام پدیدآور"][0][0:slashloc-1].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                D = FileExcelRaw["عنوان و نام پدیدآور"][0][slashloc+1:].replace('\.', '؛')
                Originators = D.split('؛')
                for O in Originators:
                    if ("مترجمان" in O):
                        Book.Translator = O[O.find("مترجمان")+len("مترجمان"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ("مترجم" in O):
                        Book.Translator = O[O.find("مترجم")+len("مترجم"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ("م\u200dت\u200dرج\u200dم" in O):
                        Book.Translator = O[O.find("م\u200dت\u200dرج\u200dم")+len("م\u200dت\u200dرج\u200dم"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ("تالیف و ترجمه" in O):
                        Book.Translator = O[O.find("تالیف و ترجمه")+len("تالیف و ترجمه"):]
                        Book.Translator = Book.Translator[Book.Translator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ("ت\u200dرج\u200dم\u200dه\u200c و ت\u200dال\u200dی\u200dف" in O):
                        Book.Translator = O[O.find("ت\u200dرج\u200dم\u200dه\u200c و ت\u200dال\u200dی\u200dف")+len("ت\u200dرج\u200dم\u200dه\u200c و ت\u200dال\u200dی\u200dف"):]
                        Book.Translator = Book.Translator[Book.Translator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ('ترجمه' in O):
                        Book.Translator = O[O.find("ترجمه")+len("ترجمه"):]
                        Book.Translator = Book.Translator[Book.Translator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ('ت\u200dرج\u200dم\u200dه' in O):
                        Book.Translator = O[O.find('ت\u200dرج\u200dم\u200dه')+len('ت\u200dرج\u200dم\u200dه'):]
                        Book.Translator = Book.Translator[Book.Translator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif (" ت\u200dرج\u200dم\u200dه\u200c و گ\u200dردآوری" in O):
                        Book.Translator = O[O.find(" ت\u200dرج\u200dم\u200dه\u200c و گ\u200dردآوری")+len(" ت\u200dرج\u200dم\u200dه\u200c و گ\u200dردآوری"):]
                        Book.Translator = Book.Translator[Book.Translator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ('ب\u200dرگ\u200dردان' in O):
                        Book.Translator = O[O.find('ب\u200dرگ\u200dردان')+len('ب\u200dرگ\u200dردان'):]
                        Book.Translator = Book.Translator[Book.Translator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ("ویراستار" in O):
                        Book.Editor = O[O.find('ویراستار')+len('ویراستار'):]
                        Book.Editor = Book.Editor[Book.Editor.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif ("وی\u200dراس\u200dت\u200dار" in O):
                        Book.Editor = O[O.find('وی\u200dراس\u200dت\u200dار')+len('وی\u200dراس\u200dت\u200dار'):]
                        Book.Editor = Book.Editor[Book.Editor.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                    elif "نویسنده و تصویرگر" in O:
                        Book.Author = O[O.find("نویسنده و تصویرگر")+len("نویسنده و تصویرگر"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        Book.Illustrator = Book.Author
                    elif "نویسنده و طراح" in O:
                        Book.Author = O[O.find("نویسنده و و طراح")+len("نویسنده و و طراح"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        Book.Illustrator = Book.Author
                    elif "ن\u200dویسنده وتصویرگر  ت" in O:
                        Book.Author = O[O.find("ن\u200dویسنده وتصویرگر  ت")+len("ن\u200dویسنده وتصویرگر  ت"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        Book.Illustrator = Book.Author
                    elif "ن\u200dوی\u200dس\u200dن\u200dده\u200c و ت\u200dص\u200dوی\u200dرگ" in O:
                        Book.Author = O[O.find("ن\u200dوی\u200dس\u200dن\u200dده\u200c و ت\u200dص\u200dوی\u200dرگ")+len("ن\u200dوی\u200dس\u200dن\u200dده\u200c و ت\u200dص\u200dوی\u200dرگ"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        Book.Author = Book.Author[Book.Author.find(" ")+1:]
                        Book.Illustrator = Book.Author
                    else:
                        if ("نویسنده" in O):
                            Book.Author = O[O.find("نویسنده")+len("نویسنده"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("نویسندگان" in O):
                            Book.Author = O[O.find("نویسندگان")+len("نویسندگان"):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("نوشته" in O):
                            # to manage نوشته‌ی or such
                            Book.Author = O[O.find("نوشته")+len("نوشته"):]
                            Book.Author = Book.Author[Book.Author.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("ن\u200dوش\u200dت\u200dه" in O):
                            Book.Author = O[O.find("ن\u200dوش\u200dت\u200dه")+len("ن\u200dوش\u200dت\u200dه"):]
                            Book.Author = Book.Author[Book.Author.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("ن\u200dویسنده" in O):
                            Book.Author = O[O.find("ن\u200dویسنده")+len("ن\u200dویسنده"):]
                            Book.Author = Book.Author[Book.Author.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("ن\u200dوی\u200dس\u200dن\u200dده" in O):
                            Book.Author = O[O.find("ن\u200dوی\u200dس\u200dن\u200dده")+len("ن\u200dوی\u200dس\u200dن\u200dده"):]
                            Book.Author = Book.Author[Book.Author.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("مولفین" in O):
                            Book.Author = O[O.find("مولفین")+len("مولفین"):]
                        if ("تصویرگر متن" in O):
                            # to manage تصویرگری or such
                            Book.Illustrator = O[O.find("تصویرگر متن")+len("تصویرگر متن"):]
                            Book.Illustrator =Book.Illustrator[Book.Illustrator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("تصویرگر" in O):
                            # to manage تصویرگری or such
                            Book.Illustrator = O[O.find("تصویرگر")+len("تصویرگر"):]
                            Book.Illustrator =Book.Illustrator[Book.Illustrator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("تصویر گر" in O):
                            # to manage تصویرگری or such
                            Book.Illustrator = O[O.find("تصویر گر")+len("تصویر گر"):]
                            Book.Illustrator =Book.Illustrator[Book.Illustrator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        elif ("ت\u200dص\u200dوی\u200dرگ\u200dر" in O):
                            Book.Illustrator = O[O.find("ت\u200dص\u200dوی\u200dرگ\u200dر")+len("ت\u200dص\u200dوی\u200dرگ\u200dر"):]
                            Book.Illustrator =Book.Illustrator[Book.Illustrator.find(" ")+1:].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                        if not ("نویسنده" in O) and not ("نوشته" in O) and not ("ن\u200dویسنده" in O) \
                            and not ("نویسندگان" in O) and not ("تصویرگر" in O) and not ("ن\u200dوی\u200dس\u200dن\u200dده" in O) \
                            and not("ت\u200dص\u200dوی\u200dرگ\u200dر" in O) and not ("تصویر گر" in O) \
                            and not ("ن\u200dوش\u200dت\u200dه" in O) and not ("تصویرگر متن" in O) \
                            and not ("مولفین" in O):
                            if len(Book.Author) == 0:
                                    Book.Author = O.strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                            else: 
                                Book.Author = Book.Author + " - " + O.strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
            elif 'نویسنده' in FileExcelRaw["عنوان و نام پدیدآور"][0]:
                Book.Name = FileExcelRaw["عنوان و نام پدیدآور"][0][0:FileExcelRaw["عنوان و نام پدیدآور"][0].find('نویسنده')].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                Book.Author = FileExcelRaw["عنوان و نام پدیدآور"][0][FileExcelRaw["عنوان و نام پدیدآور"][0].find('نویسنده')+len('نویسنده'):].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
            else:
                Book.Name = FileExcelRaw["عنوان و نام پدیدآور"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                print("Warning: نویسنده کتاب مشخص نشد -- نام: " + Book.Name)
                print("Warning: نویسنده کتاب مشخص نشد -- نام: " + Book.Name,file=f,  flush=True)
        if "مشخصات نشر" in FileExcelRaw:
            # while '\u202C\u202C' in FileExcelRaw["مشخصات نشر"][0]:
            #     FileExcelRaw["مشخصات نشر"][0] = FileExcelRaw["مشخصات نشر"][0].replace('\u202C\u202C','\u202C') 
            # while '  ' in FileExcelRaw["مشخصات نشر"][0]:
            #     FileExcelRaw["مشخصات نشر"][0] = FileExcelRaw["مشخصات نشر"][0].replace('  ',' ') 
            # while FileExcelRaw["مشخصات نشر"][0][-1] == ' ' or FileExcelRaw["مشخصات نشر"][0][-1] == '\u202C':
            #     FileExcelRaw["مشخصات نشر"][0] = FileExcelRaw["مشخصات نشر"][0][:-1]
            FileExcelRaw["مشخصات نشر"] = FileExcelRaw["مشخصات نشر"].replace('ي','ی')
            FileExcelRaw["مشخصات نشر"][0] = FileExcelRaw["مشخصات نشر"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
            if ':' in FileExcelRaw["مشخصات نشر"][0]: 
                Book.City = FileExcelRaw["مشخصات نشر"][0].split(':')[0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                PublisherInfo = FileExcelRaw["مشخصات نشر"][0][FileExcelRaw["مشخصات نشر"][0].find(':')+1:]
            else:
                PublisherInfo = FileExcelRaw["مشخصات نشر"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
            Year = PublisherInfo 
            if PublisherInfo[-1] == '.':
                Year = PublisherInfo.split('،')[-1].split('.')[0]
            elif PublisherInfo[-1] == '-':
                Year = PublisherInfo.split('،')[-1].split('-')[0]
            elif PublisherInfo[-4:].isnumeric():
                Year = PublisherInfo[-4:]
            if (unidecode(Year).strip().isnumeric()):
                Book.Publisher = PublisherInfo[0:PublisherInfo.find(Year)-1].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
                Book.Year = unidecode(Year).strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
            else:
                Book.Publisher = Year
                print("Warning: سال کتاب و ناشر از هم تفکیک نشد: " + PublisherInfo)
                print("Warning: سال کتاب و ناشر از هم تفکیک نشد: " + PublisherInfo,file=f,  flush=True)
        if "فروست" in FileExcelRaw:
            Book.Series = FileExcelRaw["فروست"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "یادداشت" in FileExcelRaw:
            Book.Notes = FileExcelRaw["یادداشت"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "موضوع" in FileExcelRaw:
            Book.Topics = FileExcelRaw["موضوع"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "شابک" in FileExcelRaw:
            Book.ISBN = '\'' + unidecode(FileExcelRaw["شابک"][0]).strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "شناسه افزوده" in FileExcelRaw:
            Book.AddedInfo = FileExcelRaw["شناسه افزوده"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "رده بندی کنگره" in FileExcelRaw:
            Book.Congress = FileExcelRaw["رده بندی کنگره"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "رده بندی دیویی" in FileExcelRaw:
            Book.Dewey = FileExcelRaw["رده بندی دیویی"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        if "شماره کتابشناسی ملی" in FileExcelRaw:
            Book.NationalNo = FileExcelRaw["شماره کتابشناسی ملی"][0].strip('\u200f\u202b\u202c\u202d\u202e\u202f .:-،')
        ResultingBooks.append(Book)
    R= pandas.DataFrame([vars(t) for t in ResultingBooks])
    R.columns = ['عنوان','نویسنده/خالق','ناشر','فروست','تصویرگر','مترجم','شابک','موضوع','ویراستار','اولین سال نشر','شهر ناشر','یادداشت','رده بندی دیویی','رده بندی کنگری','کتابشناسی ملی','شناسه افزوده']
    return R




def WriteToExcel(Book: Books, MainExcelNameAndPath: str, sheetName: str):
    BookDF = pandas.DataFrame([vars(t) for t in Book])
    BookDF.columns = ['نام','نویسنده/خالق','تصویرگر','مترجم','ویراستار','ناشر','اولین سال نشر', 'شهر ناشر', 'شابک','فروست','یادداشت','موضوع']
    if not(exists(MainExcelNameAndPath)):
        writer = pandas.ExcelWriter(MainExcelNameAndPath, engine='openpyxl', mode='w')
        BookDF.index = BookDF.index + 1 
        BookDF.to_excel(writer, sheet_name=sheetName)
    else:
        writer = pandas.ExcelWriter(MainExcelNameAndPath, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        xl = pandas.ExcelFile(MainExcelNameAndPath)
        if sheetName in xl.sheet_names:
            BookDF.index = BookDF.index + writer.sheets[sheetName].max_row
            BookDF.to_excel(writer, sheet_name=sheetName, startrow=writer.sheets[sheetName].max_row, header=None)
        else:
            BookDF.to_excel(writer, sheet_name=sheetName)
    writer.close()

def CheckBookPresenceInExcel(Book: Books, CheckSheet):
    if ('شابک' in CheckSheet.columns) and (len(Book.ISBN) >= 10):
        l = (CheckSheet['شابک'].str.find(Book.ISBN))
        l = l[l > -1].index
        if  len(l) > 0:
            return l[0],0
        else:
            return -1, 0
    else:
        return -1
                
          
        
def AppendtoExcel(BookList: Books, MainExcelNameAndPath: str, sheetName: str):
    DoWrite = 0
    if not(exists(MainExcelNameAndPath)):
        writer = pandas.ExcelWriter(MainExcelNameAndPath, engine='openpyxl', mode='w')
    else:
        try:
            writer = pandas.ExcelWriter(MainExcelNameAndPath, engine='openpyxl', mode='a')
        except:
            writer = pandas.ExcelWriter(MainExcelNameAndPath, engine='openpyxl', mode='w')
    try:
        xl = pandas.ExcelFile(MainExcelNameAndPath)
        if sheetName in xl.sheet_names: 
            SheetExists = True
        else: 
            SheetExists = False
    except:
        SheetExists = False
    if SheetExists:
        CheckSheet = pandas.read_excel(MainExcelNameAndPath, sheetName, index_col = 0)
    for Book in BookList:
        BookDF = pandas.Series(vars(Book)).to_frame().T
        print(BookDF.T)
        print(BookDF.T,file=f,  flush=True)
        if SheetExists:
            ind = CheckBookPresenceInExcel(Book, CheckSheet)
            if (ind > -1):
                key = readInput(Book.__str__()+" is present; " + CheckSheet.loc[ind].to_string() + "\r\n\tWhat should be done? (r: replace; a: add new (keep both); else escape", 'e', timeout = 20)
                if key == 'r':
                    print(" ... replacing")
                    print(" ... replacing", file=f,  flush=True)
                    DoWrite = 1
                    CheckSheet[CheckSheet['شابک'].str.contains(Book.ISBN)==True] = pandas.Series(vars(Book))
                elif key == 'a':
                    print(" ... appending")
                    print(" ... appending", file=f,  flush=True)
                    DoWrite = 1
                    BookDF.columns = CheckSheet.columns
                    BookDF.index = BookDF.index + len(CheckSheet.index) + 1
                    CheckSheet = pandas.concat([CheckSheet, BookDF])
                else:
                    print(" ... bypassing")
                    print(" ... bypassing", file=f,  flush=True)              
                    pass
            elif ind == -1:
                DoWrite = 1
                BookDF.columns = CheckSheet.columns
                BookDF.index = BookDF.index + len(CheckSheet.index) + 1
                CheckSheet = pandas.concat([CheckSheet, BookDF])
            else:
                DoWrite = 1
                CheckSheet.index = BookDF.index + len(CheckSheet.index) + 1
                try:
                    CheckSheet = pandas.concat([CheckSheet, BookDF])
                except:
                    print('Sheet ', sheetname, 'is not concatanable; Omitting sheet altogether!!')
                    print('Sheet ', sheetname, 'is not concatanable; Omitting sheet altogether!!', file=f,  flush=True)
                    CheckSheet = BookDF
                    CheckSheet.columns = ['عنوان','نویسنده/خالق','ناشر','فروست','تصویرگر','مترجم','شابک','موضوع','ویراستار','اولین سال نشر','شهر ناشر','یادداشت','رده بندی دیویی','رده بندی کنگری','کتابشناسی ملی','شناسه افزوده']
        else:
            DoWrite = 1
            CheckSheet = BookDF
            CheckSheet.columns = ['عنوان','نویسنده/خالق','ناشر','فروست','تصویرگر','مترجم','شابک','موضوع','ویراستار','اولین سال نشر','شهر ناشر','یادداشت','رده بندی دیویی','رده بندی کنگری','کتابشناسی ملی','شناسه افزوده']
            CheckSheet.index = CheckSheet.index + 1 
    if (DoWrite == 1):
        if SheetExists:
            writer.book.remove(writer.book[sheetName])
        CheckSheet.to_excel(writer, sheet_name=sheetName)
    writer.close()
        
def main():
    f = open("Log.txt", "a")
    EscapeCount = 0
    MainExcelNameAndPath = "MAIN_BOOK_LIST.xlsx"
    InputExcelNameAndPath = "کتابخانه باغ.xlsx"
    args = sys.argv[1:]
    if len(args) == 2: 
        if args[0] == '-OutExcel':
            MainExcelNameAndPath = args[1]
        elif args[0] == '-InExcel':
            InputExcelNameAndPath = args[1]
        elif args[0] == "-Escape":
            EscapeCount = int(args[1])
    elif len(args) == 4:
        if args[2] == '-OutExcel':
            MainExcelNameAndPath = args[3]
        elif args[2] == '-InExcel':
            InputExcelNameAndPath = args[3]
        elif args[2] == "-Escape":
            EscapeCount = int(args[3])
    elif len(args) == 6:
        if args[4] == '-OutExcel':
            MainExcelNameAndPath = args[5]
        elif args[4] == '-InExcel':
            InputExcelNameAndPath = args[5]
        elif args[4] == "-Escape":
            EscapeCount = int(args[5])
    
    xl = pandas.ExcelFile(InputExcelNameAndPath)
    CountAll = 0
    for sheet in xl.sheet_names:
        Book = []
        if sheet == 'کتاب های درسی' or sheet == 'کتاب در کلاس':
            continue

        BookList = pandas.read_excel(InputExcelNameAndPath, sheet_name = sheet)
        for j, bookinput in BookList.iterrows():
            CountAll = CountAll + 1 
            if CountAll <= EscapeCount:
                continue
            booktemp = Books()
            if isinstance(bookinput['نویسنده'],str):
                booktemp.Author = bookinput['نویسنده']
            if isinstance(bookinput['فروست'],str):
                booktemp.Series = bookinput['فروست']
            if isinstance(bookinput['عنوان'], str):
                booktemp.Name = bookinput['عنوان']
                if isinstance(bookinput['نشر'], str):
                    booktemp.Publisher = bookinput['نشر']
                    FileExcelRaw = SearchInNLAI(Title=bookinput['عنوان'], Publisher=bookinput['نشر'])
                    if isinstance(FileExcelRaw, int):
                        print("Retrying ...")
                        FileExcelRaw = SearchInNLAI(Title=bookinput['عنوان'], Publisher=bookinput['نشر'])
                else:
                    FileExcelRaw = SearchInNLAI(Title=bookinput['عنوان'])
                    if isinstance(FileExcelRaw, int):
                        print("Retrying ...")
                        FileExcelRaw = SearchInNLAI(Title=bookinput['عنوان'])
            else:
                print("Book Name must be present")
                print("Book Name must be present",file=f,  flush=True)
                continue
            print(j, CountAll, file=f,  flush=True)
            if FileExcelRaw is None:
                AppendtoExcel([booktemp], MainExcelNameAndPath, sheet)
            elif isinstance(FileExcelRaw, int):
                AppendtoExcel([booktemp], MainExcelNameAndPath, sheet)
            else:
                Book = AnalyzeExcelRaw(FileExcelRaw)
                AppendtoExcel([Book], MainExcelNameAndPath, sheet)

        #WriteToExcel(Book, MainExcelNameAndPath, sheet)
    f.close()

if __name__ == "__main__":
    main()
    