import re
import unicodedata
from typing import Optional, Tuple, List

# Cadastro canônico dos integrantes autorizados da Bancada STI
BANCADA_MEMBROS = [
    {
        "nome": "Paulo Henrique Gonçalves Rezende",
        "primeiro_nome": "Paulo",
        "telefone": "5567992471379",
        "telefone_formatado": "+55 67 99247-1379",
        "aliases": ["paulo henrique", "paulo rezende", "paulo", "rezende"]
    },
    {
        "nome": "Reginaldo da Silva Bandeira",
        "primeiro_nome": "Reginaldo",
        "telefone": "5567991455446",
        "telefone_formatado": "+55 67 99145-5446",
        "aliases": ["reginaldo da silva bandeira", "reginaldo bandeira", "reginaldo da silva", "reginaldo", "bandeira"]
    },
    {
        "nome": "Luiz Leonardo Villalba",
        "primeiro_nome": "Luiz",
        "telefone": "5567996477799",
        "telefone_formatado": "+55 67 99647-7799",
        "aliases": ["luiz leonardo villalba", "luiz villalba", "luiz leonardo", "luiz", "villalba"]
    }
]

def _normalize(s: str) -> str:
    """Remove acentos, pontuação e converte para minúsculas para comparação flexível."""
    if not s or not isinstance(s, str):
        return ""
    # Normalização unicode (NFD separa acentos das letras)
    n = unicodedata.normalize("NFD", s.strip().lower())
    n = "".join(c for c in n if unicodedata.category(c) != "Mn")
    n = re.sub(r"[^a-z0-9\s]", " ", n)
    return re.sub(r"\s+", " ", n).strip()

def resolve_bancada_member(text: str) -> Optional[dict]:
    """
    Identifica se uma string de texto/nome corresponde a um dos 3 membros autorizados da Bancada.
    Suporta nomes parciais como 'Paulo Rezende', 'Reginaldo Bandeira', 'Luiz Villalba'.
    Retorna o dicionário completo do membro (ou None se não for da bancada).
    """
    if not text:
        return None
        
    norm_text = _normalize(text)
    if not norm_text:
        return None

    # Varre os membros em ordem de especificidade
    for membro in BANCADA_MEMBROS:
        norm_nome = _normalize(membro["nome"])
        # Correspondência exata do nome completo
        if norm_nome == norm_text or norm_nome in norm_text:
            return membro

        # Correspondência por aliases
        for alias in membro["aliases"]:
            norm_alias = _normalize(alias)
            # Verifica correspondência de palavras inteiras usando regex
            pattern = rf"\b{re.escape(norm_alias)}\b"
            if re.search(pattern, norm_text):
                return membro

    return None

def extract_all_bancada_members(text: str) -> List[dict]:
    """
    Dada uma lista de nomes ou string com separadores (ex: 'Paulo Rezende / Thyago',
    'Reginaldo Bandeira e Luiz Villalba'), identifica todos os membros da bancada presentes.
    """
    if not text:
        return []

    # Divide por separadores comuns: /, virgula, ponto e virgula, ' e '
    partes = re.split(r"[/,;]|\s+e\s+", str(text))
    encontrados = []
    telefones_vistos = set()

    for parte in partes:
        parte_limpa = parte.strip()
        if not parte_limpa:
            continue
        m = resolve_bancada_member(parte_limpa)
        if m and m["telefone"] not in telefones_vistos:
            encontrados.append(m)
            telefones_vistos.add(m["telefone"])

    return encontrados
