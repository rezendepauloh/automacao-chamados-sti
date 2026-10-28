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
from datetime import datetime

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
            if isinstance(key, tuple):
                idx, col = key
                val = self.df.data.get(col, {})
                if isinstance(val, (list, dict)):
                    return val[idx]
                return val
            elif isinstance(key, int):
                row_dict = {col: self.df.data[col][key] for col in self.df.columns if col in self.df.data and key < len(self.df.data[col])}
                return MockSeries(row_dict)
            return None

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
            self.empty = len(self.data) == 0 or (len(self.data) > 0 and len(next(iter(self.data.values()))) == 0)
            self.at = _MockAtIndexer(self)
            self.loc = _MockAtIndexer(self)
            self.iloc = _MockAtIndexer(self)

        def head(self, n=5):
            if not self.data:
                return self.copy()
            new_data = {col: self.data[col][:n] for col in self.columns if col in self.data}
            return MockDataFrame(data=new_data, columns=list(self.columns))

        def groupby(self, by_col):
            if not self.data or by_col not in self.data:
                return []
            groups = {}
            rows_count = len(self)
            for i in range(rows_count):
                k = self.data[by_col][i]
                groups.setdefault(k, []).append(i)
            result = []
            for k, idxs in groups.items():
                g_data = {col: [self.data[col][i] for i in idxs] for col in self.columns if col in self.data}
                result.append((k, MockDataFrame(data=g_data, columns=list(self.columns))))
            return result

        def copy(self):
            return MockDataFrame(data=dict(self.data), columns=list(self.columns))

        def dropna(self, subset=None):
            if not self.data:
                return self.copy()
            rows_count = len(next(iter(self.data.values())))
            cols_to_check = subset if subset else self.columns
            valid_indices = []
            for i in range(rows_count):
                keep = True
                for col in cols_to_check:
                    val = self.data.get(col, [None] * rows_count)[i]
                    if val is None or (isinstance(val, float) and val != val) or val == "":
                        keep = False
                        break
                if keep:
                    valid_indices.append(i)
            new_data = {col: [self.data[col][i] for i in valid_indices] for col in self.columns}
            return MockDataFrame(data=new_data, columns=list(self.columns))

        def to_dict(self, orient="records"):
            if not self.data:
                return [] if orient == "records" else {}
            rows_count = len(next(iter(self.data.values())))
            if orient == "records":
                return [{col: self.data[col][i] for col in self.columns} for i in range(rows_count)]
            return dict(self.data)

        def iterrows(self):
            if isinstance(self.data, dict) and self.data:
                rows_count = len(next(iter(self.data.values())))
                for i in range(rows_count):
                    row_dict = {col: self.data[col][i] for col in self.columns if i < len(self.data[col])}
                    yield i, MockSeries(row_dict)

        def to_excel(self, *args, **kwargs):
            return True

        def to_csv(self, *args, **kwargs):
            return ""

        def __getitem__(self, item):
            if isinstance(self.data, dict) and item in self.data:
                col_data = self.data[item]
                return _MockColumn(col_data if isinstance(col_data, list) else [col_data])
            return _MockColumn([])

        def __setitem__(self, item, value):
            if isinstance(value, _MockColumn):
                self.data[item] = list(value.values)
            elif isinstance(value, list):
                self.data[item] = list(value)
            else:
                rows_count = len(self) if self.data else 1
                self.data[item] = [value] * rows_count
            if item not in self.columns:
                self.columns.append(item)
            self.empty = len(self.data) == 0 or len(self) == 0

        def __len__(self):
            if isinstance(self.data, dict) and self.data:
                return len(next(iter(self.data.values())))
            return 0

    class _MockColumn:
        def __init__(self, values):
            self.values = list(values)

        def tolist(self):
            return list(self.values)

        def dropna(self):
            return _MockColumn([v for v in self.values if v is not None and not (isinstance(v, float) and v != v)])

        def fillna(self, fill_val=""):
            return _MockColumn([fill_val if (v is None or (isinstance(v, float) and v != v)) else v for v in self.values])

        def astype(self, _t):
            return _MockColumn([str(v) if _t is str else v for v in self.values])

        def replace(self, pat, repl, regex=False):
            import re
            if regex:
                return _MockColumn([re.sub(pat, repl, str(v)) for v in self.values])
            return _MockColumn([repl if v == pat else v for v in self.values])

        def apply(self, fn):
            return _MockColumn([fn(v) for v in self.values])

        def sum(self):
            return sum(int(v) for v in self.values if v is not None)

        @property
        def str(self):
            col_self = self
            class _StrMethods:
                def strip(self_s):
                    return _MockColumn([str(v).strip() for v in col_self.values])
            return _StrMethods()

        def __len__(self):
            return len(self.values)

        def __iter__(self):
            return iter(self.values)

    class MockSeries(dict):
        def __init__(self, data=None, **kwargs):
            if isinstance(data, dict):
                super().__init__(data)
            elif isinstance(data, (list, tuple)):
                super().__init__({i: v for i, v in enumerate(data)})
            else:
                super().__init__(**kwargs)

        def to_dict(self):
            return dict(self)

        def __getattr__(self, item):
            if item in self:
                return self[item]
            raise AttributeError(f"'MockSeries' object has no attribute '{item}'")

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
    mock_pd.Timestamp = datetime
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
        def _wrap(fn):
            fn.clear = MagicMock()
            return fn
        if len(args) == 1 and callable(args[0]) and not kwargs:
            return _wrap(args[0])
        return _wrap

    _decorator_helper.clear = MagicMock()

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

# Mock openpyxl
try:
    import openpyxl
except ImportError:
    mock_openpyxl = types.ModuleType("openpyxl")
    mock_openpyxl.Workbook = MagicMock()
    mock_openpyxl.load_workbook = MagicMock()
    mock_styles = types.ModuleType("openpyxl.styles")
    mock_styles.PatternFill = MagicMock()
    mock_styles.Font = MagicMock()
    mock_styles.Alignment = MagicMock()
    mock_styles.Border = MagicMock()
    mock_styles.Side = MagicMock()
    mock_openpyxl.styles = mock_styles
    sys.modules["openpyxl"] = mock_openpyxl
    sys.modules["openpyxl.styles"] = mock_styles

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
class _MockLDAPConnection:
    def __init__(self, *args, **kwargs):
        self.bound = False
        self.result = {"description": "Invalid credentials"}
    def bind(self):
        return False
    def unbind(self):
        self.bound = False

_mock_module_if_missing("ldap3", {
    "Server": MagicMock(),
    "Connection": _MockLDAPConnection,
    "ALL": "ALL",
    "SUBTREE": "SUBTREE",
    "NTLM": "NTLM",
    "SIMPLE": "SIMPLE",
})
_mock_module_if_missing("ldap3.core", {})
_mock_module_if_missing("ldap3.core.exceptions", {
    "LDAPException": Exception,
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

# 7b. Mock de itsdangerous e selenium caso ausentes no host
try:
    import itsdangerous
except ImportError:
    import json
    import base64
    import hmac
    import hashlib
    import time as _time

    class _SignatureExpired(Exception):
        pass

    class _BadSignature(Exception):
        pass

    class _URLSafeTimedSerializer:
        def __init__(self, secret_key, salt="default"):
            self.secret_key = str(secret_key).encode()
            self.salt = str(salt).encode()

        def dumps(self, obj):
            payload = {"d": obj, "t": _time.time()}
            raw = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode()
            sig = hmac.new(self.secret_key + self.salt, raw.encode(), hashlib.sha256).hexdigest()
            return f"{raw}.{sig}"

        def loads(self, token, max_age=None):
            if not isinstance(token, str) or "." not in token:
                raise _BadSignature("Invalid token format")
            raw, sig = token.rsplit(".", 1)
            expected = hmac.new(self.secret_key + self.salt, raw.encode(), hashlib.sha256).hexdigest()
            if not hmac.compare_digest(sig, expected):
                raise _BadSignature("Bad signature")
            try:
                payload = json.loads(base64.urlsafe_b64decode(raw.encode()).decode())
            except Exception as e:
                raise _BadSignature("Corrupted payload") from e
            if max_age is not None and (_time.time() - payload.get("t", 0)) > max_age:
                raise _SignatureExpired("Token expired")
            return payload.get("d")

    _mock_module_if_missing("itsdangerous", {
        "URLSafeTimedSerializer": _URLSafeTimedSerializer,
        "SignatureExpired": _SignatureExpired,
        "BadSignature": _BadSignature,
    })

_mock_module_if_missing("selenium", {"webdriver": MagicMock()})
_mock_module_if_missing("selenium.webdriver", {"Chrome": MagicMock(), "ChromeOptions": MagicMock()})
_mock_module_if_missing("selenium.webdriver.common", {})
_mock_module_if_missing("selenium.webdriver.common.by", {"By": MagicMock()})
_mock_module_if_missing("selenium.webdriver.common.keys", {"Keys": MagicMock()})
_mock_module_if_missing("selenium.webdriver.chrome", {})
_mock_module_if_missing("selenium.webdriver.chrome.options", {"Options": MagicMock()})
_mock_module_if_missing("selenium.webdriver.chrome.service", {"Service": MagicMock()})
_mock_module_if_missing("selenium.webdriver.support", {})
_mock_module_if_missing("selenium.webdriver.support.ui", {"WebDriverWait": MagicMock()})
_mock_module_if_missing("selenium.webdriver.support.expected_conditions", {})
_mock_module_if_missing("selenium.common", {})
_mock_module_if_missing("selenium.common.exceptions", {
    "TimeoutException": Exception,
    "WebDriverException": Exception,
    "NoSuchElementException": Exception,
})
_mock_module_if_missing("webdriver_manager", {})
_mock_module_if_missing("webdriver_manager.chrome", {"ChromeDriverManager": MagicMock()})

import unicodedata as _unicodedata
_mock_module_if_missing("unidecode", {
    "unidecode": lambda s: "".join(
        c for c in _unicodedata.normalize("NFD", str(s)) if _unicodedata.category(c) != "Mn"
    ),
})
_mock_module_if_missing("pdfplumber", {
    "open": MagicMock(),
})
_mock_module_if_missing("playwright", {})
_mock_module_if_missing("playwright.sync_api", {"sync_playwright": MagicMock()})
_mock_module_if_missing("playwright.async_api", {"async_playwright": MagicMock()})

# 8. Mock de bs4 (BeautifulSoup)
try:
    import bs4
except ImportError:
    import re
    import html as html_lib

    class MockTag:
        def __init__(self, name, attrs=None, **kwargs):
            self.name = name
            self.attrs = attrs or {}
            self.attrs.update(kwargs)

        def get(self, key, default=None):
            return self.attrs.get(key, default)

        def __getitem__(self, key):
            return self.attrs[key]

        def __setitem__(self, key, value):
            self.attrs[key] = value

        def decompose(self):
            pass

        def replace_with(self, new_tag):
            pass

        def find_all(self, *args, **kwargs):
            return []

        def __str__(self):
            attrs_str = " ".join(f'{k}="{v}"' for k, v in self.attrs.items())
            return f"<{self.name} {attrs_str} />" if attrs_str else f"<{self.name} />"

    class MockSoup:
        def __init__(self, html="", parser=None):
            self.html = str(html) if html else ""

        def new_tag(self, name, **kwargs):
            attrs = kwargs.get("attrs", {})
            for k, v in kwargs.items():
                if k != "attrs":
                    attrs[k] = v
            return MockTag(name, attrs=attrs)

        def find_all(self, *args, **kwargs):
            # Procura tags básicas baseadas no nome ou predicado
            if args:
                target = args[0]
                # Suporta lista de tags como ["i", "span"]
                tag_names = target if isinstance(target, (list, tuple)) else [target]

                if any(t == "i" for t in tag_names):
                    matches = []
                    for m in re.finditer(r'<i(\s+[^>]*)?>.*?</i>|<i(\s+[^>]*)?/>', self.html, flags=re.DOTALL | re.IGNORECASE):
                        attrs_raw = m.group(1) or ""
                        attrs = {}
                        for attr_match in re.finditer(r'([a-zA-Z0-9_-]+)=[\'"]([^\'"]*)[\'"]', attrs_raw):
                            attrs[attr_match.group(1)] = html_lib.unescape(attr_match.group(2))
                        class MockI(MockTag):
                            def __init__(self, full, attrs_dict, soup_ref):
                                super().__init__("i", attrs_dict)
                                self.full = full
                                self.soup_ref = soup_ref
                            def decompose(self):
                                self.soup_ref.html = self.soup_ref.html.replace(self.full, "")
                        matches.append(MockI(m.group(0), attrs, self))
                    if target == "i" or (isinstance(target, (list, tuple)) and set(target) == {"i"}):
                        return matches
                    if isinstance(target, (list, tuple)):
                        return matches
                
                if "div" in tag_names and kwargs.get("class_") == "imagePlugin":
                    matches = []
                    pattern = re.compile(r'<div([^>]*class=[\'"][^\'"]*imagePlugin[^\'"]*[\'"][^>]*)>(.*?)</div>', re.DOTALL | re.IGNORECASE)
                    for m in pattern.finditer(self.html):
                        attrs_raw = m.group(1)
                        full = m.group(0)
                        attrs = {}
                        for attr_match in re.finditer(r'([a-zA-Z0-9_-]+)=[\'"]([^\'"]*)[\'"]', attrs_raw):
                            attrs[attr_match.group(1)] = html_lib.unescape(attr_match.group(2))
                        
                        class MockDiv(MockTag):
                            def __init__(self, full_str, attrs_dict, soup_ref):
                                super().__init__("div", attrs_dict)
                                self.full_str = full_str
                                self.soup_ref = soup_ref
                            def replace_with(self, new_content):
                                self.soup_ref.html = self.soup_ref.html.replace(self.full_str, str(new_content))

                        matches.append(MockDiv(full, attrs, self))
                    return matches

                if "div" in tag_names and kwargs.get("class_") == "sp-video-embed":
                    matches = []
                    pattern = re.compile(r'<div([^>]*class=[\'"][^\'"]*sp-video-embed[^\'"]*[\'"][^>]*)>(.*?)</div>', re.DOTALL | re.IGNORECASE)
                    for m in pattern.finditer(self.html):
                        attrs_raw = m.group(1)
                        full = m.group(0)
                        attrs = {}
                        for attr_match in re.finditer(r'([a-zA-Z0-9_-]+)=[\'"]([^\'"]*)[\'"]', attrs_raw):
                            attrs[attr_match.group(1)] = html_lib.unescape(attr_match.group(2))
                        
                        class MockVideoDiv(MockTag):
                            def __init__(self, full_str, attrs_dict, soup_ref):
                                super().__init__("div", attrs_dict)
                                self.full_str = full_str
                                self.soup_ref = soup_ref
                            def replace_with(self, new_content):
                                self.soup_ref.html = self.soup_ref.html.replace(self.full_str, str(new_content))

                        matches.append(MockVideoDiv(full, attrs, self))
                    return matches
                
                if "img" in tag_names:
                    matches = []
                    for m in re.finditer(r'<img(\s+[^>]*)?>', self.html, flags=re.IGNORECASE):
                        attrs_raw = m.group(1) or ""
                        full = m.group(0)
                        attrs = {}
                        for attr_match in re.finditer(r'([a-zA-Z0-9_-]+)=[\'"]([^\'"]*)[\'"]', attrs_raw):
                            attrs[attr_match.group(1)] = html_lib.unescape(attr_match.group(2))
                        
                        class MockImg(MockTag):
                            def __init__(self, full_str, attrs_dict, soup_ref):
                                super().__init__("img", attrs_dict)
                                self.full_str = full_str
                                self.soup_ref = soup_ref
                            def decompose(self):
                                self.soup_ref.html = self.soup_ref.html.replace(self.full_str, "")
                        matches.append(MockImg(full, attrs, self))
                    return matches

            # Se for função de predicado (como controldata_divs)
            if args and callable(args[0]):
                pred = args[0]
                matches = []
                pattern = re.compile(r'<div([^>]*controldata[^>]*)>(.*?)</div>', re.DOTALL | re.IGNORECASE)
                for m in pattern.finditer(self.html):
                    attrs_raw = m.group(1)
                    full = m.group(0)
                    attrs = {}
                    # Captura tanto 'attr="valor"' quanto 'attr=\'valor\''
                    for attr_match in re.finditer(r'([a-zA-Z0-9_-]+)=(?:"([^"]*)"|\'([^\']*)\')', attrs_raw):
                        val = attr_match.group(2) if attr_match.group(2) is not None else attr_match.group(3)
                        attrs[attr_match.group(1)] = html_lib.unescape(val)
                    
                    class MockPredDiv(MockTag):
                        def __init__(self, full_str, attrs_dict, soup_ref):
                            super().__init__("div", attrs_dict)
                            self.full_str = full_str
                            self.soup_ref = soup_ref
                        def replace_with(self, new_content):
                            self.soup_ref.html = self.soup_ref.html.replace(self.full_str, str(new_content))

                    tag_obj = MockPredDiv(full, attrs, self)
                    if pred(tag_obj):
                        matches.append(tag_obj)
                return matches

            return []

        def __str__(self):
            return self.html

    mock_bs4 = types.ModuleType("bs4")
    mock_bs4.BeautifulSoup = MockSoup
    sys.modules["bs4"] = mock_bs4

