import streamlit as st
import tkinter
from tkinter import filedialog
from pathlib import Path
import LibraryTools
import shutil
import pandas
from BookSearchLib import (Books, CheckBookPresenceInExcel, SearchInNLAI, AnalyzeExcelRaw)
#	src:	url('Yekan.eot'); /* IE9 Compat Modes */
#	src:	url('Yekan.eot?#iefix') format('embedded-opentype'), /* IE6-IE8 */
#			url('Yekan.woff2') format('woff2'), /* Modern Browsers */
#			url('Yekan.woff') format('woff'), /* Modern Browsers */
#			url('Yekan.otf') format('opentype'), /* Open Type Font */	
#			url('Yekan.ttf') format('truetype'); /* Safari, Android, iOS */

st.markdown(
        """
        <style>
@import url('https://v1.fontapi.ir/css/Yekan');
@font-face {
	font-family: 'Yekan';
    src: local('Tekan.woff2'), url('assets/fonts/Yekan.woff2') format('woff2');
	font-weight: normal;
	font-style: normal;
	text-rendering: optimizeLegibility;
	font-display: auto;
}
    div.element-container div.row-widget.stRadio[role="radiogroup"] {
        color : purple
    }
    html, body, [class*="css"]  {
    font-family: 'Yekan';
    text-align: center;
    direction: rtl;
    }
    </style>

    """,
        unsafe_allow_html=True,
    )

if 'ExcelDetermined' in st.session_state:
    del st.session_state['ExcelDetermined']
st.markdown("<div id='linkto_top'></div>", unsafe_allow_html=True)    
st.title("جستجوی کتاب")
if "BookInfo_Loaded" not in st.session_state:
    SearchBook = Books()
    BookDF = None
    col1, col2 = st.columns([1, 12])
    check = False
    with col2:
        with st.form("BookSearchForm", clear_on_submit=False):
            SearchBook.ISBN = st.text_input("شابک")
            SearchBook.Name = st.text_input("نام")
            SearchBook.Author = st.text_input("نویسنده/خالق/تصویرگر ...")
            SearchBook.Publisher = st.text_input("ناشر")
            SearchBook.Topics = st.text_input("موضوع")
            st.text("در کدام صفحات بگردم؟")
            neededcheckboxes = len(st.session_state['XL'])
            if neededcheckboxes < 5:
                cols = st.columns(neededcheckboxes)
            elif (neededcheckboxes % 3) == 0:
                cols = st.columns(3)
                Mod = 3
            else:
                cols = st.columns(4)
                Mod = 4
            sheetselect = [False] * neededcheckboxes
            for i, sheetName in enumerate(st.session_state['XL']):
                with cols[i % Mod]:
                    sheetselect[i] = st.checkbox(sheetName, value=True)
            if SearchBook.Author != "" or SearchBook.Publisher != "" or SearchBook.Name != "" or SearchBook.ISBN != "" or SearchBook.Topics != "":
                BookDF = pandas.Series(vars(SearchBook)).to_frame().T
                BookDF.columns = ['عنوان','نویسنده/خالق','ناشر','فروست','تصویرگر','مترجم','شابک','موضوع','ویراستار','اولین سال نشر','شهر ناشر','یادداشت','رده بندی دیویی','رده بندی کنگری','کتابشناسی ملی','شناسه افزوده']
            if st.form_submit_button("بگرد"):
                if BookDF is not None:
                    for i,sheet in enumerate(st.session_state['XL']):
                        if sheetselect[i]:
                            CheckSheet = st.session_state['Sheets'][i]
                            if ('شابک' in CheckSheet.columns) and (len(SearchBook.ISBN) >= 10):
                                indfind = (CheckSheet['شابک'].str.find(SearchBook.ISBN))
                                indfind = indfind[indfind > -1].index
                                if  len(indfind) > 0:
                                    CheckSheet = CheckSheet.loc[indfind]
                                    check = True
                            else:
                                for i,(l,k) in enumerate(vars(SearchBook).items()):
                                    if len(k) > 1:
                                        indfind = CheckSheet[BookDF.columns[i]].str.find(k)
                                        indfind = indfind[indfind > -1].index
                                        if l == 'Author':
                                            indfind2 = CheckSheet['مترجم'].str.find(k)
                                            indfind2 = indfind2[indfind > -1].index
                                            if  len(indfind2) > 0:
                                                indfind.append(indfind2)
                                            indfind2 = CheckSheet['تصویرگر'].str.find(k)
                                            indfind2 = indfind2[indfind > -1].index
                                            if  len(indfind2) > 0:
                                                indfind.append(indfind2)
                                        if  len(indfind) > 0:
                                            CheckSheet = CheckSheet.loc[indfind]                                
                                            check = True
                                        else:
                                            break
    if check:
        st.title("اینها رو پیدا کردم")
        st.table(CheckSheet.T)
    else:
        st.subheader("چنین کتابی پیدا نشد. در اینترنت بگردم؟")
        NLAI = Books()
        if st.button("بگرد"):
            if len(SearchBook.ISBN) <= 9:
                NLAI.ISBN = None
            else:
                NLAI.ISBN = SearchBook.ISBN
            if len(SearchBook.Name) <= 1:
                NLAI.Name = None
            else:
                NLAI.Name = SearchBook.Name
            if len(SearchBook.Author) <= 1:
                NLAI.Author = None
            else:
                NLAI.Author = SearchBook.Author
            if len(SearchBook.Publisher) <= 1:
                NLAI.Publisher = None
            else:
                NLAI.Publisher = SearchBook.Publisher
            FileExcelRaw = SearchInNLAI(ISBN=NLAI.ISBN, Author=NLAI.Author, Publisher=NLAI.Publisher, Title=NLAI.Name)
            if isinstance(FileExcelRaw, int):
                st.text("دارم تلاش می‌کنم")
                FileExcelRaw = SearchInNLAI(ISBN=SearchBook.ISBN, Author=SearchBook.Author, Publisher=SearchBook.Publisher, Title=SearchBook.Name)
            if isinstance(FileExcelRaw, int):
                if FileExcelRaw == -2:
                    st.subheader("موفق نبود. اینترنت را چک کنید یا بعداً تلاش کنید")
                elif FileExcelRaw == -1:
                    st.subheader("این جستجو در اینترنت هم بی‌پاسخ بود! از درستی اطلاعات کتاب مطمئنید؟")
            else:
                InternetFoundBooks = AnalyzeExcelRaw(FileExcelRaw)                    
                st.table(InternetFoundBooks)


            

