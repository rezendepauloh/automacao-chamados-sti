import streamlit as st
import pandas as pd

st.title("Test Data Editor Selection")

@st.dialog("Detalhes do Chamado")
def show_details(row):
    st.write(f"Detalhes do chamado {row['A']}")
    if st.button("Fechar"):
        st.rerun()

df = pd.DataFrame({"A": [1, 2, 3], "B": ["x", "y", "z"]})

# Test data editor
result = st.data_editor(
    df,
    key="my_editor",
    on_select="rerun",
    selection_mode="single-row"
)

st.write("Result type:", type(result))
st.write("Result:", result)

# Print selection from session state
if "my_editor" in st.session_state:
    st.write("Session State:", st.session_state["my_editor"])
