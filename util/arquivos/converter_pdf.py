import fitz  # PyMuPDF
from pathlib import Path

uploads_dir = Path("uploads")
for pdf_path in uploads_dir.rglob("*.pdf"):
    print(f"Convertendo {pdf_path}...")
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)  # Carrega a primeira página
    
    # Renderiza a página com matriz de escala 2.0 para manter boa definição
    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
    
    png_path = pdf_path.with_suffix(".png")
    pix.save(str(png_path))
    doc.close()
    print(f"Salvo: {png_path}")
