# -*- coding: utf-8 -*-
"""
Helper universal de testes para o Sistema Bancada.
Garante mocks seguros e transparentes de dependências opcionais (pandas, streamlit, keyring, ldap3, schedule, etc.)
permitindo que a suíte inteira rode com 100% de confiabilidade e isolamento tanto no host local
quanto dentro do contêiner Docker.
"""

import sys
import types
from unittest.mock import MagicMock

def _mock_module_if_missing(name, attrs=None):
    if name not in sys.modules:
        try:
            __import__(name)
        except ImportError:
            mod = types.ModuleType(name)
            if attrs:
                for k, v in attrs.items():
                    setattr(mod, k, v)
            sys.modules[name] = mod
            return mod
    return sys.modules[name]

# 1. Mock seguro de pandas se não estiver instalado
try:
    import pandas as pd
except ImportError:
    mock_pd = types.ModuleType("pandas")

    class _MockAtIndexer:
        def __init__(self, df):
            self.df = df

        def __setitem__(self, key, value):
            idx, col = key
            if col not in self.df.data:
                self.df.data[col] = {}
            if isinstance(self.df.data[col], list):
                while len(self.df.data[col]) <= idx:
                    self.df.data[col].append(None)
                self.df.data[col][idx] = value
            elif isinstance(self.df.data[col], dict):
                self.df.data[col][idx] = value

        def __getitem__(self, key):
            idx, col = key
            val = self.df.data.get(col, {})
            if isinstance(val, (list, dict)):
                return val[idx]
            return val

    class MockDataFrame:
        def __init__(self, data=None, columns=None, *args, **kwargs):
            if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                cols = list(data[0].keys()) if not columns else list(columns)
                d_dict = {c: [item.get(c) for item in data] for c in cols}
                self.data = d_dict
                self.columns = cols
            elif isinstance(data, dict):
                self.data = dict(data)
                self.columns = columns or list(self.data.keys())
            else:
                self.data = {}
                self.columns = columns or []
            self.empty = len(self.data) == 0
            self.at = _MockAtIndexer(self)
            self.loc = _MockAtIndexer(self)
            self.iloc = _MockAtIndexer(self)

        def copy(self):
            return MockDataFrame(data=dict(self.data), columns=list(self.columns))

        def iterrows(self):
            if isinstance(self.data, dict) and self.data:
                rows_count = len(next(iter(self.data.values())))
                for i in range(rows_count):
                    row_dict = {col: self.data[col][i] for col in self.columns if i < len(self.data[col])}
                    yield i, row_dict

        def to_excel(self, *args, **kwargs):
            return True

        def to_csv(self, *args, **kwargs):
            return ""

        def __getitem__(self, item):
            if isinstance(self.data, dict) and item in self.data:
                col_data = self.data[item]
                mock_col = MagicMock()
                mock_col.values = col_data if isinstance(col_data, list) else [col_data]
                mock_col.tolist = lambda: col_data if isinstance(col_data, list) else [col_data]
                return mock_col
            mock_col = MagicMock()
            mock_col.values = []
            mock_col.tolist = lambda: []
            return mock_col

        def __len__(self):
            if isinstance(self.data, dict) and self.data:
                return len(next(iter(self.data.values())))
            return 0

    class MockSeries(MagicMock):
        pass

    def _mock_read_sql_query(sql, con, params=None):
        try:
            cur = con.cursor()
            if params:
                cur.execute(sql, params)
            else:
                cur.execute(sql)
            rows = cur.fetchall()
            col_names = [d[0] for d in cur.description] if cur.description else []
            data_dict = {col: [r[i] for r in rows] for i, col in enumerate(col_names)}
            return MockDataFrame(data=data_dict, columns=col_names)
        except Exception:
            return MockDataFrame()

    mock_pd.DataFrame = MockDataFrame
    mock_pd.Series = MockSeries
    mock_pd.isna = lambda x: x is None or (isinstance(x, float) and x != x)
    mock_pd.notna = lambda x: not (x is None or (isinstance(x, float) and x != x))
    mock_pd.read_excel = MagicMock(return_value=MockDataFrame())
    mock_pd.read_csv = MagicMock(return_value=MockDataFrame())
    mock_pd.read_sql_query = _mock_read_sql_query
    mock_pd.to_datetime = MagicMock(side_effect=lambda x: x)
    sys.modules["pandas"] = mock_pd

# Mock sklearn e submódulos
try:
    import sklearn
except ImportError:
    mock_sklearn = types.ModuleType("sklearn")
    sys.modules["sklearn"] = mock_sklearn

for submod, attrs in [
    ("sklearn.pipeline", {"Pipeline": MagicMock()}),
    ("sklearn.svm", {"LinearSVC": MagicMock()}),
    ("sklearn.feature_extraction", {}),
    ("sklearn.feature_extraction.text", {"TfidfVectorizer": MagicMock()}),
    ("sklearn.model_selection", {"StratifiedKFold": MagicMock(), "GridSearchCV": MagicMock()}),
    ("sklearn.naive_bayes", {"MultinomialNB": MagicMock(), "ComplementNB": MagicMock()}),
    ("sklearn.ensemble", {"RandomForestClassifier": MagicMock()}),
    ("sklearn.metrics", {"classification_report": MagicMock(return_value="")}),
]:
    m = types.ModuleType(submod)
    for k, v in attrs.items():
        setattr(m, k, v)
    sys.modules[submod] = m
    # Vincula ao módulo pai
    parts = submod.split(".")
    if len(parts) == 2:
        setattr(sys.modules[parts[0]], parts[1], m)
    elif len(parts) == 3:
        if parts[1] in dir(sys.modules[parts[0]]):
            setattr(getattr(sys.modules[parts[0]], parts[1]), parts[2], m)

# 2. Mock seguro de streamlit
try:
    import streamlit as st
except ImportError:
    mock_st = types.ModuleType("streamlit")
    mock_st.session_state = {}
    mock_st.query_params = {}
    
    def _decorator_helper(*args, **kwargs):
        if len(args) == 1 and callable(args[0]) and not kwargs:
            return args[0]
        return lambda fn: fn

    mock_st.dialog = _decorator_helper
    mock_st.cache_data = _decorator_helper
    mock_st.cache_resource = _decorator_helper
    
    mock_comp = types.ModuleType("streamlit.components")
    mock_comp_v1 = types.ModuleType("streamlit.components.v1")
    mock_comp_v1.html = MagicMock()
    mock_comp.v1 = mock_comp_v1
    mock_st.components = mock_comp

    sys.modules["streamlit"] = mock_st
    sys.modules["streamlit.components"] = mock_comp
    sys.modules["streamlit.components.v1"] = mock_comp_v1

# Mock joblib e spacy
try:
    import joblib
except ImportError:
    mock_joblib = types.ModuleType("joblib")
    mock_joblib.load = MagicMock(return_value=MagicMock())
    mock_joblib.dump = MagicMock()
    sys.modules["joblib"] = mock_joblib

try:
    import spacy
except ImportError:
    mock_spacy = types.ModuleType("spacy")
    mock_nlp = MagicMock()
    mock_doc = MagicMock()
    mock_sents = [MagicMock(text="Bom dia"), MagicMock(text="Computador travando")]
    mock_doc.sents = mock_sents
    mock_nlp.return_value = mock_doc
    mock_spacy.load = MagicMock(return_value=mock_nlp)
    sys.modules["spacy"] = mock_spacy

    mock_lang = types.ModuleType("spacy.lang")
    mock_pt = types.ModuleType("spacy.lang.pt")
    mock_sw = types.ModuleType("spacy.lang.pt.stop_words")
    mock_sw.STOP_WORDS = {"de", "a", "o", "que", "e", "do", "da", "em", "um", "para"}
    mock_pt.stop_words = mock_sw
    mock_lang.pt = mock_pt
    mock_spacy.lang = mock_lang

    sys.modules["spacy.lang"] = mock_lang
    sys.modules["spacy.lang.pt"] = mock_pt
    sys.modules["spacy.lang.pt.stop_words"] = mock_sw

# Mock fuzzywuzzy
try:
    import fuzzywuzzy
except ImportError:
    mock_fuzzy = types.ModuleType("fuzzywuzzy")
    mock_proc = types.ModuleType("fuzzywuzzy.process")
    mock_proc.extractOne = MagicMock(return_value=("Campo Grande", 90))
    mock_fuzzy.process = mock_proc
    sys.modules["fuzzywuzzy"] = mock_fuzzy
    sys.modules["fuzzywuzzy.process"] = mock_proc

# Mock xlsxwriter
try:
    import xlsxwriter
except ImportError:
    mock_xlsx = types.ModuleType("xlsxwriter")
    mock_wb_mod = types.ModuleType("xlsxwriter.workbook")
    mock_wb_mod.Workbook = MagicMock()
    mock_xlsx.Workbook = mock_wb_mod.Workbook
    sys.modules["xlsxwriter"] = mock_xlsx
    sys.modules["xlsxwriter.workbook"] = mock_wb_mod

# Mock dotenv
try:
    import dotenv
except ImportError:
    mock_dotenv = types.ModuleType("dotenv")
    mock_dotenv.load_dotenv = MagicMock()
    mock_dotenv.find_dotenv = MagicMock(return_value="")
    sys.modules["dotenv"] = mock_dotenv

# 3. Mock de keyring (cofre de senhas)
try:
    import keyring
except ImportError:
    mock_keyring = types.ModuleType("keyring")
    _in_memory_keys = {}
    mock_keyring.get_password = lambda service, username: _in_memory_keys.get(f"{service}:{username}")
    mock_keyring.set_password = lambda service, username, pwd: _in_memory_keys.update({f"{service}:{username}": pwd})
    mock_keyring.delete_password = lambda service, username: _in_memory_keys.pop(f"{service}:{username}", None)
    sys.modules["keyring"] = mock_keyring

# 4. Mock de schedule
_mock_module_if_missing("schedule", {
    "every": MagicMock(return_value=MagicMock()),
    "run_pending": MagicMock(),
    "get_jobs": MagicMock(return_value=[]),
    "clear": MagicMock(),
})

# 5. Mock de ldap3
_mock_module_if_missing("ldap3", {
    "Server": MagicMock(),
    "Connection": MagicMock(),
    "ALL": "ALL",
    "SUBTREE": "SUBTREE",
    "NTLM": "NTLM",
})

# 6. Mock de icalendar
_mock_module_if_missing("icalendar", {
    "Calendar": MagicMock(),
    "Event": MagicMock(),
})

# 7. Mock de psycopg2
_mock_module_if_missing("psycopg2", {
    "connect": MagicMock(),
})
_mock_module_if_missing("psycopg2.extras", {
    "RealDictCursor": MagicMock(),
})
