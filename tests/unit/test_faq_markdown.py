# -*- coding: utf-8 -*-
"""
Testes unitários para o serviço de Manuais Markdown (.md) do SharePoint (src/services/faq_markdown.py)
e para a persistência/substituição dos antigos Site Pages na tabela faqs.
"""
import sqlite3
import tempfile
import unittest
from pathlib import Path

from src.services.faq_markdown import (
    DEFAULT_MARKDOWN_SERVER_REL_PATH,
    DEFAULT_SP_SITE_URL,
    deduplicate_markdown_faqs,
    extract_markdown_media_references,
    extract_markdown_title,
    extract_sharepoint_folder_rel_path,
    extract_sharepoint_site_url,
    format_category_name,
    is_html_faq_content,
    render_markdown_for_streamlit,
    resolve_markdown_media_url,
    scan_local_markdown_faqs,
)


class TestFaqMarkdownService(unittest.TestCase):
    def test_extract_sharepoint_folder_rel_path_from_allitems_url(self):
        url = (
            "https://ministeriopublicoms.sharepoint.com/sites/dit-manutencao/"
            "Shared%20Documents/Forms/AllItems.aspx?"
            "id=%2Fsites%2Fdit%2Dmanutencao%2FShared%20Documents%2FTutoriais%2DFAQs%2FManuais%20Markdown"
            "&viewid=fa6fa6a3%2Dc035%2D4588%2Da4c8%2D6df7e448e42d"
        )
        rel = extract_sharepoint_folder_rel_path(url)
        self.assertEqual(rel, DEFAULT_MARKDOWN_SERVER_REL_PATH)
        self.assertEqual(extract_sharepoint_site_url(url), DEFAULT_SP_SITE_URL)

    def test_extract_sharepoint_folder_rel_path_fallback(self):
        self.assertEqual(extract_sharepoint_folder_rel_path(""), DEFAULT_MARKDOWN_SERVER_REL_PATH)
        self.assertEqual(
            extract_sharepoint_folder_rel_path(r"C:\Users\teste\OneDrive\Manuais Markdown"),
            DEFAULT_MARKDOWN_SERVER_REL_PATH,
        )

    def test_format_category_name(self):
        self.assertEqual(format_category_name("formatacao"), "Formatação")
        self.assertEqual(format_category_name("telefonia_e_oxe"), "Telefonia e OXE")
        self.assertEqual(format_category_name("geral"), "Geral")
        self.assertEqual(format_category_name("rede_e_vpn"), "Rede e Vpn")

    def test_extract_markdown_title(self):
        md = "# 💿 Formatação na Bancada\n\nTexto introdutório."
        self.assertEqual(extract_markdown_title(md, "formatacao_na_bancada.md"), "Formatação na Bancada")

        md_bracket = "# 📘 [TOKEN] As diferenças do Safe Net\n\nConteúdo."
        self.assertEqual(
            extract_markdown_title(md_bracket, "token.md"),
            "[TOKEN] As diferenças do Safe Net",
        )

        self.assertEqual(extract_markdown_title("Sem heading", "guia_de_remessa.md"), "Guia de Remessa")

    def test_deduplicate_markdown_faqs_prefers_thematic_over_geral(self):
        items = [
            {
                "filename": "formatacao_na_bancada.md",
                "folder": "geral",
                "titulo": "Formatação na Bancada",
                "tipo_faq": "Geral",
                "url": "https://sp/geral/formatacao_na_bancada.md",
                "conteudo": "Versão antiga",
            },
            {
                "filename": "formatacao_na_bancada.md",
                "folder": "formatacao",
                "titulo": "Formatação na Bancada",
                "tipo_faq": "Formatação",
                "url": "https://sp/formatacao/formatacao_na_bancada.md",
                "conteudo": "Versão refatorada completa",
            },
            {
                "filename": "somente_geral.md",
                "folder": "geral",
                "titulo": "Manual Exclusivo do Geral",
                "tipo_faq": "Geral",
                "url": "https://sp/geral/somente_geral.md",
                "conteudo": "Conteúdo único",
            },
        ]
        deduped = deduplicate_markdown_faqs(items)
        self.assertEqual(len(deduped), 2)
        urls = {it["url"] for it in deduped}
        self.assertIn("https://sp/formatacao/formatacao_na_bancada.md", urls)
        self.assertIn("https://sp/geral/somente_geral.md", urls)
        self.assertNotIn("https://sp/geral/formatacao_na_bancada.md", urls)

    def test_deduplicate_by_normalized_title_across_geral(self):
        items = [
            {
                "filename": "protocolar_um_documento_no_saj.md",
                "folder": "geral",
                "titulo": "Protocolar documento no SAJ",
                "tipo_faq": "Geral",
                "url": "https://sp/geral/protocolar_um_documento_no_saj.md",
                "conteudo": "Antigo",
            },
            {
                "filename": "protocolar_documento_saj.md",
                "folder": "administrativo",
                "titulo": "Protocolar documento no SAJ",
                "tipo_faq": "Administrativo",
                "url": "https://sp/administrativo/protocolar_documento_saj.md",
                "conteudo": "Novo",
            },
        ]
        deduped = deduplicate_markdown_faqs(items)
        self.assertEqual(len(deduped), 1)
        self.assertEqual(deduped[0]["tipo_faq"], "Administrativo")

    def test_resolve_markdown_media_url_and_path_traversal_guard(self):
        md_rel = "/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Manuais Markdown/formatacao/formatacao_na_bancada.md"
        resolved_img = resolve_markdown_media_url("<../../Imagens FAQ/Formatação/SCCM 01.png>", md_rel)
        self.assertEqual(
            resolved_img,
            "https://ministeriopublicoms.sharepoint.com/sites/dit-manutencao/Shared%20Documents/Tutoriais-FAQs/Imagens%20FAQ/Formata%C3%A7%C3%A3o/SCCM%2001.png",
        )

        resolved_vid = resolve_markdown_media_url(
            "<../../Vídeos FAQ/Impressora/Como descobrir.mp4>",
            md_rel,
        )
        self.assertIn("/V%C3%ADdeos%20FAQ/Impressora/Como%20descobrir.mp4", resolved_vid)

        # Path Traversal fora de /sites/ deve ser bloqueado
        self.assertEqual(resolve_markdown_media_url("../../../../../../etc/passwd", md_rel), "")

        # URL externa preservada
        self.assertEqual(
            resolve_markdown_media_url("https://example.com/img.png", md_rel),
            "https://example.com/img.png",
        )

    def test_extract_markdown_media_references(self):
        md = (
            "![Tela 1](<../../Imagens FAQ/Formatação/SCCM 01.png>)\n"
            "[Vídeo](<../../Vídeos FAQ/Impressora/demo.mp4>)\n"
            "[Link externo](https://example.com)\n"
        )
        md_rel = "/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Manuais Markdown/formatacao/teste.md"
        imgs, vids = extract_markdown_media_references(md, md_rel)
        self.assertEqual(len(imgs), 1)
        self.assertEqual(len(vids), 1)
        self.assertIn("SCCM%2001.png", imgs[0])
        self.assertIn("demo.mp4", vids[0])

    def test_render_markdown_for_streamlit(self):
        md = (
            "# 💿 Formatação na Bancada\n\n"
            "> [!TIP]\n"
            "> Use a aba **Limpeza AD**.\n\n"
            "![Tela SCCM](<../../Imagens FAQ/Formatação/SCCM 01.png>)\n\n"
            "[Assistir Demo](<../../Vídeos FAQ/Impressora/demo.mp4>)\n"
        )
        rendered = render_markdown_for_streamlit(
            md,
            faq_title="Formatação na Bancada",
            md_server_rel_url="/sites/dit-manutencao/Shared Documents/Tutoriais-FAQs/Manuais Markdown/formatacao/formatacao_na_bancada.md",
            image_loader=lambda url, title: "data:image/png;base64,AAAA",
            video_loader=lambda url, title: ("", None),
        )
        # H1 duplicado removido
        self.assertNotIn("# 💿 Formatação na Bancada", rendered)
        # Callout convertido
        self.assertIn("💡 Dica", rendered)
        self.assertIn("<strong>Limpeza AD</strong>", rendered)
        # Imagem embutida em base64 com legenda
        self.assertIn('src="data:image/png;base64,AAAA"', rendered)
        self.assertIn("Tela SCCM", rendered)
        # Card de vídeo renderizado com link para o SharePoint
        self.assertIn("Assistir Demo", rendered)
        self.assertIn("Assistir no SharePoint", rendered)

    def test_is_html_faq_content(self):
        self.assertTrue(is_html_faq_content('<div data-sp-controldata="{}"><p>Oi</p></div>'))
        self.assertFalse(is_html_faq_content("# Título\n\n> [!NOTE]\n> Texto em Markdown"))

    def test_scan_local_markdown_faqs_and_db_replacement(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base = Path(tmpdir)
            (base / "README.md").write_text("# Índice", encoding="utf-8")
            (base / "geral").mkdir()
            (base / "formatacao").mkdir()
            (base / "geral" / "manual_a.md").write_text("# 📘 Manual A\nAntigo", encoding="utf-8")
            (base / "formatacao" / "manual_a.md").write_text("# 💿 Manual A\nRefatorado", encoding="utf-8")
            (base / "geral" / "manual_b.md").write_text("# 📘 Manual B\nÚnico", encoding="utf-8")

            items = scan_local_markdown_faqs(base)
            self.assertEqual(len(items), 2)
            cats = {it["titulo"]: it["tipo_faq"] for it in items}
            self.assertEqual(cats["Manual A"], "Formatação")
            self.assertEqual(cats["Manual B"], "Geral")

            # Testa substituição completa dos antigos Site Pages (.aspx) na tabela faqs
            from src.services.faq_markdown import persist_markdown_faqs_to_db

            conn = sqlite3.connect(":memory:")
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE faqs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    titulo TEXT NOT NULL,
                    tipo_faq TEXT NOT NULL,
                    url TEXT NOT NULL UNIQUE,
                    conteudo TEXT,
                    data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                "INSERT INTO faqs (titulo, tipo_faq, url, conteudo) VALUES (?, ?, ?, ?)",
                ("Página Antiga", "Tutoriais SharePoint", "https://sp/SitePages/Antiga.aspx", "<div>Old</div>"),
            )
            conn.commit()

            count = persist_markdown_faqs_to_db(cursor, conn, items)
            self.assertEqual(count, 2)

            rows = cursor.execute("SELECT titulo, tipo_faq, url FROM faqs").fetchall()
            self.assertEqual(len(rows), 2)
            self.assertFalse(any(r[2].endswith(".aspx") for r in rows))
            conn.close()

    def test_mirror_markdown_faqs_to_disk_preserves_subdirs_and_cleans_stale(self):
        from src.services.faq_markdown import mirror_markdown_faqs_to_disk

        with tempfile.TemporaryDirectory() as tmpdir:
            base = Path(tmpdir)
            # Cria arquivo obsoleto que deve ser removido no espelhamento
            stale_dir = base / "geral"
            stale_dir.mkdir(parents=True)
            stale_file = stale_dir / "antigo_duplicado.md"
            stale_file.write_text("# Antigo", encoding="utf-8")

            items = [
                {
                    "filename": "formatacao_na_bancada.md",
                    "folder": "formatacao",
                    "titulo": "Formatação na Bancada",
                    "tipo_faq": "Formatação",
                    "url": "https://sp/formatacao/formatacao_na_bancada.md",
                    "conteudo": "# 💿 Formatação na Bancada\nConteúdo",
                },
                {
                    "filename": "manual_impressora.md",
                    "folder": "impressoras",
                    "titulo": "Manual Impressora",
                    "tipo_faq": "Impressoras",
                    "url": "https://sp/impressoras/manual_impressora.md",
                    "conteudo": "# Impressora\nConteúdo",
                },
            ]

            written = mirror_markdown_faqs_to_disk(base, items)
            self.assertEqual(written, 2)
            self.assertTrue((base / "formatacao" / "formatacao_na_bancada.md").exists())
            self.assertTrue((base / "impressoras" / "manual_impressora.md").exists())
            self.assertFalse(stale_file.exists())
            self.assertFalse(stale_dir.exists())


if __name__ == "__main__":
    unittest.main()
