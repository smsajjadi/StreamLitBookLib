import streamlit as st
import tkinter
from tkinter import filedialog
from pathlib import Path
import LibraryTools
import shutil
import pandas
from LibraryTools import (
    delete_page,
    add_page, 
    nav_page
)
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
    html, body, [class*="css"]  {
    font-family: 'Yekan';
    text-align: center;
    direction: rtl;
    }
    </style>

    """,
        unsafe_allow_html=True,
    )

@st.cache_data
def loadExcel(ExcelFileName):
    xl = pandas.ExcelFile(ExcelFileName)
    CheckSheet = []
    for i,SheetName in enumerate(xl.sheet_names):
        CheckSheet.append(pandas.read_excel(ExcelFileName, SheetName, index_col = 0))
    return CheckSheet 

@st.cache_data
def loadxl(ExcelFileName):
    xl = pandas.ExcelFile(ExcelFileName)
    return xl.sheet_names

def move_font_files():
    STREAMLIT_STATIC_PATH = Path(st.__path__[0]) / "static"
    CSS_PATH = STREAMLIT_STATIC_PATH / "assets/fonts/"
    if not CSS_PATH.is_dir():
        CSS_PATH.mkdir()

    css_file = CSS_PATH / "Yekan.woff2"
    if not css_file.exists():
        shutil.copy("Yekan.woff2", css_file)

delete_page('streamlittest.py', 'SearchBook')
move_font_files()
#tkniter for folder selection -- if needed
root = tkinter.Tk()
root.withdraw()
root.wm_attributes('-topmost', 1)

if 'Buildup' not in st.session_state:
    st.session_state['Buildup'] = 0
if 'SelectBaseFile' not in st.session_state:
    st.session_state['SelectBaseFile'] = 0
if 'ConfigExcelPresent' not in st.session_state:
    st.session_state['ConfigExcelPresent'] = False

st.text(st.session_state['Buildup'])
if st.session_state['Buildup'] == 1:
    st.write('Please select a folder to build-up the excel in:')
    clicked = st.button('Folder Picker')
    if clicked:
        dirname = st.text_input('Selected folder:', filedialog.askdirectory(master=root))
        if len(dirname) > 0:
            try:
                st.session_state['Buildup'] = 0
                MainExcel = dirname+'MainBookList.xlsx'
                writer = pandas.ExcelWriter(MainExcel, engine='openpyxl', mode='w')
                st.session_state['BasicExcelFile'] = MainExcel
                st.text_input("Filename", "MainBookList.xlsx")
                st.button("OK")
            except:
                pass
elif st.session_state['SelectBaseFile'] == 1:
    MainExcel = st.file_uploader("فایل مورد نظر را انتخاب کنید")
    if MainExcel is not None:
        try:
            writer = pandas.ExcelWriter(MainExcel.name, engine='openpyxl', mode='a')
            st.text(MainExcel.name)
        except:
            st.session_state['CountState'] = -2
            #st.session_state['Buildup'] = 1
        st.session_state['BasicExcelFile'] = MainExcel.name
        st.session_state['ExcelDetermined'] = 1
        st.text(st.session_state['BasicExcelFile'])
        if st.button("فایل شما را در مسیر خود کپی کنم؟"):
            f = open("./Library.Config","w")
            f.write("MainBookList.xlsx")
            f.close()
            f = open("./MainBookList.xlsx", "wb")
            f.write(MainExcel.getbuffer())
            f.close()
            if 'ExcelDetermined' in st.session_state:
                del st.session_state['ExcelDetermined']
            st.session_state['SelectBaseFile'] = 0
            st.experimental_rerun()


else:
    if 'ExcelDetermined' not in st.session_state:
        st.session_state['Buildup'] = 0
        st.title("برنامه کتابخانه شخصی")
        st.subheader("اکسل پایه کتابخانه")
        col1, col2, col3 = st.columns([2, 3, 1])
        with col2:
            try:
                f = open("Library.Config","r")
                XlsName=f.readline()
                f.close()
                st.text(XlsName)
                writer = pandas.ExcelWriter(XlsName, engine='openpyxl', mode='a')
                st.session_state['ConfigExcelPresent'] = True
                #except:
                #    st.text("Could not open file")
            except:
                st.text('file open exception')
            if st.session_state['ConfigExcelPresent']:
                if st.button("**پیش‌فرض**            (MainBookList.xlsx) ",use_container_width=True):
                    MainExcel = "MainBookList.xlsx"
    #                try:
                    writer = pandas.ExcelWriter(MainExcel, engine='openpyxl', mode='a')
                    st.session_state['BasicExcelFile'] = MainExcel
                    st.session_state['ExcelDetermined'] = 1
                    add_page('streamlittest.py', 'SearchBook')
    #                    del st.session_state['ExcelDetermined']
                    st.text(st.session_state['ExcelDetermined'])
                    st.session_state['Sheets'] = loadExcel(st.session_state['BasicExcelFile'])
                    st.session_state['XL'] = loadxl(st.session_state['BasicExcelFile'])
                    nav_page('SearchBook')
    #                except:
    #                    st.text('MainBookList.xlsx پیدا نشد -- از پایه بسازیم؟')
    #                    st.session_state['Buildup'] = 1
    #                    st.button("از پایه بسازیم")
            if st.button("بگذار بیابم",use_container_width=True):
                st.session_state['SelectBaseFile'] = 1
                st.experimental_rerun()
            elif st.button("از پایه می‌سازم",use_container_width=True):
                st.session_state['Buildup'] = 1
                st.experimental_rerun()
    else:
        pass            
