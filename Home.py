import streamlit as st
from PIL import Image
import os

st.set_page_config(
    page_title="Portale Strumenti Contabili per Amazon",
    page_icon="🏢",
    layout="centered"
)

# Mostra il logo se presente
logo_path = "logo.png"
if os.path.exists(logo_path):
    st.image(Image.open(logo_path), width=300)

st.title("Benvenuto nel Portale Strumenti Contabili per Amazon")
st.markdown("---")

st.markdown("""
### 👈 Usa il menù laterale per selezionare un'app.

Abbiamo riunito i calcolatori in un unico portale sicuro:

* **📄 PDF Extractor:** Trascina o Carica l'Estratto Conto Amazon in PDF per mappare i valori e salvarli nel file Excel AM_.._.._Estratto Conto Tab.
* **📊 Riclassicatore Estratto Conto Amazon:** Trascina o Carica l'Estratto Conto Tab per ottenere la versione riclassificata ed i corrispettivi per l'Italia o l'Estero.

---
💡 *Nota sulla privacy: I file che carichi vengono elaborati temporaneamente dalla memoria del tuo browser e distrutti immediatamente dopo l'uso. Nessun dato finanziario viene salvato su internet.*
""")
# --- FOOTER ---
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    """
    <div style='text-align: center; color: #888; font-size: 14px; margin-top: 20px;'>
        ⚡ Powered by <b>iannovins</b>
    </div>
    """, 
    unsafe_allow_html=True
)
