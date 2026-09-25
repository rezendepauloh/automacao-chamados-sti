import sqlite3
from pathlib import Path

# Localiza o arquivo chamados.db na raiz do projeto ou diretório atual
BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "chamados.db"
if not DB_PATH.exists():
    DB_PATH = Path("chamados.db")

def limpar_localidades():
    if not DB_PATH.exists():
        print(f"❌ Banco de dados não encontrado em: {DB_PATH.resolve()}")
        return

    print(f"🔧 Conectando ao banco: {DB_PATH.resolve()}")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # 1. Remove chamados espúrios com ID 'nan' ou nulo
    cursor.execute("DELETE FROM comentarios WHERE chamado_id = 'nan' OR chamado_id IS NULL")
    comentarios_nan = cursor.rowcount
    cursor.execute("DELETE FROM chamados WHERE id = 'nan' OR id IS NULL")
    chamados_nan = cursor.rowcount

    # 2. Limpar cidade_predio
    cursor.execute("""
    UPDATE chamados 
    SET cidade_predio = '' 
    WHERE LOWER(TRIM(cidade_predio)) IN ('nan', 'none', 'null', '<na>') OR cidade_predio IS NULL
    """)
    cp_cleaned = cursor.rowcount

    # 3. Limpar unidade
    cursor.execute("""
    UPDATE chamados 
    SET unidade = '' 
    WHERE LOWER(TRIM(unidade)) IN ('nan', 'none', 'null', '<na>') 
       OR LOWER(unidade) LIKE '%não encontrad% no ad%' 
       OR LOWER(unidade) LIKE '%nao encontrad% no ad%'
       OR unidade IS NULL
    """)
    un_cleaned = cursor.rowcount

    # 4. Limpar localidade_fisica com nan ou erro de AD
    cursor.execute("""
    UPDATE chamados 
    SET localidade_fisica = 'Não identificada'
    WHERE LOWER(localidade_fisica) LIKE '%nan%' 
       OR LOWER(localidade_fisica) LIKE '%não encontrad%' 
       OR LOWER(localidade_fisica) LIKE '%nao encontrad%'
       OR LOWER(TRIM(localidade_fisica)) IN ('none', 'null', '<na>', '', 'n/d')
       OR localidade_fisica IS NULL
    """)
    loc_cleaned = cursor.rowcount

    # 5. Para chamados onde cidade_predio está preenchido, define localidade_fisica como a Cidade / Prédio (sem ' - Sede' e sem unidade interna)
    cursor.execute("""
    UPDATE chamados
    SET localidade_fisica = TRIM(REPLACE(cidade_predio, ' - Sede', ''))
    WHERE cidade_predio IS NOT NULL AND TRIM(cidade_predio) != ''
      AND (
        localidade_fisica = 'Não identificada'
        OR localidade_fisica LIKE '% - %ª PJ%'
        OR localidade_fisica LIKE '% - %º PJ%'
        OR localidade_fisica LIKE '% - Promotoria%'
        OR localidade_fisica LIKE '% - Procuradoria%'
      )
    """)
    reconst_cp = cursor.rowcount

    cursor.execute("""
    UPDATE chamados
    SET localidade_fisica = TRIM(unidade)
    WHERE localidade_fisica = 'Não identificada'
      AND (cidade_predio IS NULL OR TRIM(cidade_predio) = '')
      AND unidade IS NOT NULL AND TRIM(unidade) != ''
    """)
    reconst_un = cursor.rowcount

    # 6. Padronizar prédios de Campo Grande sem sufixos de setores internos
    cursor.execute("""
    UPDATE chamados
    SET localidade_fisica = 'Campo Grande - PGJ'
    WHERE localidade_fisica LIKE 'Campo Grande - PGJ - %'
    """)
    cg_pgj_cleaned = cursor.rowcount

    cursor.execute("""
    UPDATE chamados
    SET localidade_fisica = 'Campo Grande - DMP'
    WHERE localidade_fisica LIKE 'Campo Grande - DMP - %'
    """)
    cg_dmp_cleaned = cursor.rowcount

    cursor.execute("""
    UPDATE chamados
    SET localidade_fisica = 'Campo Grande - Rua da Paz'
    WHERE localidade_fisica LIKE 'Campo Grande - Rua da Paz - %'
    """)
    cg_rp_cleaned = cursor.rowcount

    cursor.execute("""
    UPDATE chamados
    SET localidade_fisica = 'Campo Grande - Chácara Cachoeira'
    WHERE localidade_fisica LIKE 'Campo Grande - Chácara Cachoeira - %'
    """)
    cg_cc_cleaned = cursor.rowcount

    # 7. Gerar títulos inteligentes para chamados sem título (ex: CitSmart)
    try:
        from src.tag_classifier import generate_synthetic_title
    except ImportError:
        try:
            from tag_classifier import generate_synthetic_title
        except ImportError:
            generate_synthetic_title = None

    titles_generated = 0
    if generate_synthetic_title:
        cursor.execute("SELECT id, tag, descricao FROM chamados WHERE titulo IS NULL OR TRIM(titulo) = '' OR LOWER(TRIM(titulo)) IN ('none', 'nan', 'null', 'sem título', '<na>')")
        rows_to_title = cursor.fetchall()
        for r_id, r_tag, r_desc in rows_to_title:
            new_title = generate_synthetic_title(r_tag, r_desc)
            cursor.execute("UPDATE chamados SET titulo = ? WHERE id = ?", (new_title, r_id))
            titles_generated += 1

    conn.commit()
    conn.close()

    print("✅ Limpeza e padronização concluída com sucesso:")
    print(f"   - Chamados / Comentários espúrios removidos: {chamados_nan} / {comentarios_nan}")
    print(f"   - Campos 'cidade_predio' corrigidos: {cp_cleaned}")
    print(f"   - Campos 'unidade' corrigidos: {un_cleaned}")
    print(f"   - Títulos gerados por IA/NLP para chamados vazios: {titles_generated}")
    print(f"   - 'localidade_fisica' com 'nan' ou erro do AD corrigidos: {loc_cleaned}")
    print(f"   - Localidades padronizadas (Cidade / Prédio): {reconst_cp}")
    print(f"   - Localidades recuperadas (Apenas Unidade): {reconst_un}")
    print(f"   - Localidades Campo Grande padronizadas: {cg_pgj_cleaned + cg_dmp_cleaned + cg_rp_cleaned + cg_cc_cleaned}")

if __name__ == "__main__":
    limpar_localidades()
