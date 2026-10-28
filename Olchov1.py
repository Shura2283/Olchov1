Python 3.12.10 (tags/v3.12.10:0cc8128, Apr  8 2025, 12:21:36) [MSC v.1943 64 bit (AMD64)] on win32
Enter "help" below or click "Help" above for more information.
>>> import streamlit as st
... import streamlit.components.v1 as components
... 
... # Sahifa sozlamalari
... st.set_page_config(
...     page_title="O'lchov Birliklari Ta'lim Platformasi",
...     page_icon="📐",
...     layout="wide"
... )
... 
... # HTML faylingiz tarkibi
... html_content = """
... <!-- SHU YERGA O'LCHOV..TXT ICHIDAGI HTML KODNI TO'LIQ QO'YASIZ -->
... """
... 
... # HTML komponentini butun ekranda ko'rsatish
