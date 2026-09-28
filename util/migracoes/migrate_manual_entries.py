import sys
from pathlib import Path

# Adiciona a raiz do projeto e a pasta src ao sys.path
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / "src"))

from src.manual_entries import get_manual_entries
from src.database import add_unidade_manual, get_unidades_manuais


def migrate_manual_entries():
    """Migra todos os dicionários de manual_entries.py para a tabela unidades_manuais do SQLite."""
    entries = get_manual_entries()
    print(f"Encontrados {len(entries)} registros manuais em manual_entries.py. Iniciando migração...")

    count = 0
    for item in entries:
        cidade = item.get("Cidade", item.get("cidade", ""))
        tipo = item.get("Tipo", item.get("tipo", ""))
        setor = item.get("Setor", item.get("setor", ""))
        sigla = item.get("Sigla", item.get("sigla", ""))
        titular = item.get("Titular", item.get("titular", ""))
        unidade_predio = item.get("Unidade (Prédio)", item.get("unidade_predio", ""))
        telefone = item.get("Telefone", item.get("telefone", ""))
        url = item.get("URL", item.get("url", ""))

        if setor or cidade:
            add_unidade_manual(
                cidade=cidade,
                tipo=tipo,
                setor=setor,
                sigla=sigla,
                titular=titular,
                unidade_predio=unidade_predio,
                telefone=telefone,
                url=url
            )
            count += 1

    df_db = get_unidades_manuais()
    print(f"Migração concluída com sucesso! Total de {count} entradas processadas. Registros no SQLite: {len(df_db)}")


if __name__ == "__main__":
    migrate_manual_entries()
