# -*- coding: utf-8 -*-
import sqlite3
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd

from config import (
    INPUT_DIR_BRUTOS,
    OUTPUT_DIR_TRATADOS,
    DEBUG_DIR_PREPROCESS,
    setup_logging,
    save_df_to_excel_formatted,
    cleanup_old_files
)
from terminal import print_header, CYAN

# ---------------------------
# Logging
# ---------------------------
logger = setup_logging(DEBUG_DIR_PREPROCESS / "preprocess_oxe.log", __name__)


def salvar_banco_dados(df: pd.DataFrame):
    """Salva/Atualiza a tabela 'central_telefonica' no banco de dados SQLite chamados.db preservando todas as colunas."""
    try:
        db_path = Path("chamados.db")
        conn = sqlite3.connect(db_path)
        
        df_db = df.copy()
        # Normalização de nomes de colunas para padrão de banco SQLite
        col_rename = {
            c: c.lower()
                .replace(" / ", "_")
                .replace("/", "_")
                .replace(" ", "_")
                .replace(".", "")
                .replace("ç", "c")
                .replace("ã", "a")
                .replace("é", "e")
            for c in df_db.columns
        }
        df_db.rename(columns=col_rename, inplace=True)
        
        df_db.to_sql("central_telefonica", conn, if_exists="replace", index=False)
        conn.close()
        logger.info(f"✅ Tabela 'central_telefonica' atualizada no banco de dados SQLite chamados.db ({len(df)} registros, {len(df_db.columns)} colunas).")
    except Exception as db_err:
        logger.error(f"Erro ao salvar dados da Central Telefônica no SQLite: {db_err}")


def preprocess_oxe() -> bool:
    """
    Pré-processa e normaliza os dados brutos da Central Telefônica OXE.
    """
    print_header("PRÉ-PROCESSAMENTO OXE - CENTRAL TELEFÔNICA", color=CYAN)
    logger.info("=== Iniciando Pré-processamento dos dados da Central Telefônica (OXE) ===")

    files = sorted(INPUT_DIR_BRUTOS.glob("Central_Telefonica_OXE_*.xlsx"), key=lambda f: f.stat().st_mtime)
    if not files:
        logger.error("Nenhum arquivo bruto 'Central_Telefonica_OXE_*.xlsx' encontrado em 01 - Dados Brutos.")
        return False

    latest_file = files[-1]
    logger.info(f"Processando arquivo bruto mais recente: {latest_file.name}")

    try:
        df = pd.read_excel(latest_file, dtype=str)
        df.fillna("", inplace=True)
    except Exception as read_err:
        logger.error(f"Erro ao ler arquivo Excel bruto '{latest_file.name}': {read_err}")
        return False

    if df.empty:
        logger.warning("O arquivo bruto selecionado está vazio.")
        return False

    registros_tratados = []

    for _, row in df.iterrows():
        ramal = str(row.get("Ramal", "") or row.get("Directory_Number", "")).strip()
        if not ramal or ramal.lower() == "nan":
            continue

        name = str(row.get("Nome / Titular", "") or row.get("Annu_Name", "")).strip()
        comp = str(row.get("Complemento", "") or row.get("Annu_First_Name", "")).strip()

        # Nome exibido unificado
        if name and comp and comp.lower() != name.lower():
            nome_exibido = f"{name} {comp}".strip()
        elif name:
            nome_exibido = name
        elif comp:
            nome_exibido = comp
        else:
            nome_exibido = "Não informado"

        tipo_estacao = str(row.get("Tipo de Estação", "") or row.get("Station_Type", "")).strip()
        subtipo = str(row.get("Subtipo", "") or row.get("Station_Sub_Type", "")).strip()

        ip = str(row.get("Endereço IP", "") or row.get("IP_Address", "")).strip()
        mac = str(row.get("MAC Address", "") or row.get("Ethernet_Address", "")).strip().upper()

        rack = str(row.get("Rack", "") or row.get("Equipment_Address_Rack", "")).strip()
        board = str(row.get("Placa", "") or row.get("Equipment_Address_Board", "")).strip()
        terminal = str(row.get("Terminal", "") or row.get("Equipment_Address_Terminal", "")).strip()

        # Formatação do endereço físico de hardware (Rack/Board/Terminal)
        if rack not in ("", "-", "255") and board not in ("", "-", "255"):
            endereco_hardware = f"R:{rack} | P:{board} | T:{terminal}"
        else:
            endereco_hardware = "-"

        # Classificação do Tipo/Categoria do Dispositivo
        if mac and mac != "-":
            categoria = "Telefone IP Físico"
            status_equip = "Ativo com Hardware (MAC)"
        elif ip and ip != "-":
            categoria = "Telefone IP / Softphone"
            status_equip = "Ativo com IP"
        elif "ANALOG" in tipo_estacao.upper():
            categoria = "Ramal Analógico"
            status_equip = "Sem IP / MAC"
        elif "VIRTUAL" in tipo_estacao.upper():
            categoria = "Ramal Virtual"
            status_equip = "Sem IP / MAC"
        else:
            categoria = "Ramal Genérico / Outros"
            status_equip = "Sem IP / MAC"

        # Copia todos os atributos brutos originais do Excel e adiciona/sobrescreve com os calculados
        rec = dict(row)
        rec["Ramal"] = ramal
        rec["Nome Exibido"] = nome_exibido
        rec["Endereço Equipamento"] = endereco_hardware
        rec["Categoria Dispositivo"] = categoria
        rec["Status Equipamento"] = status_equip
        rec["MAC Address"] = mac if mac else "-"
        rec["Endereço IP"] = ip if ip else "-"
        
        registros_tratados.append(rec)

    df_tratado = pd.DataFrame(registros_tratados)


    # 2. Exporta para 02 - Dados tratados
    ts = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    out_file = OUTPUT_DIR_TRATADOS / f"Central_Telefonica_OXE_Tratados_{ts}.xlsx"

    widths = {
        "Ramal": 12,
        "Nome / Titular": 25,
        "Complemento": 25,
        "Nome Exibido": 30,
        "Tipo de Estação": 25,
        "Subtipo": 18,
        "Função / Role": 20,
        "Grupo de Captura": 18,
        "Cat. Rede Pública": 18,
        "Login Externo": 18,
        "E-mail": 30,
        "Rack": 8,
        "Placa": 8,
        "Terminal": 8,
        "Endereço Equipamento": 20,
        "Endereço IP": 18,
        "MAC Address": 20,
        "Categoria Dispositivo": 25,
        "Status Equipamento": 25
    }


    save_df_to_excel_formatted(
        df_tratado, out_file, sheet_name="Central Telefônica",
        widths=widths
    )

    logger.info(f"✅ Arquivo tratado gerado com sucesso: {out_file.name} ({len(df_tratado)} registros)")

    # Limpeza de backups antigos (mantém 10 recentes)
    cleanup_old_files(OUTPUT_DIR_TRATADOS, "Central_Telefonica_OXE_Tratados_*.xlsx", keep_count=10)

    # 3. Salva no banco de dados SQLite chamados.db
    salvar_banco_dados(df_tratado)

    return True


if __name__ == "__main__":
    import sys
    if preprocess_oxe():
        sys.exit(0)
    else:
        sys.exit(1)
