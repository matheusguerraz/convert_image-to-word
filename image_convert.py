from PIL import Image
import pytesseract
from docx import Document
import os

# Caminho para o executável do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\usuario\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

def image_to_text(image_path):
    # Abre a imagem
    img = Image.open(image_path)

    # Usa o Tesseract para extrair o texto da imagem
    text = pytesseract.image_to_string(img)

    return text

def create_word_document(text, output_path='output.docx'):
    # Cria um documento do Word
    doc = Document()

    # Adiciona o texto ao documento
    doc.add_paragraph(text)

    # Salva o documento
    doc.save(output_path)

if __name__ == "__main__":
    # Obtém o diretório do script
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Obtém a lista de arquivos na pasta do script
    image_files = [f for f in os.listdir(script_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]

    # Itera sobre todas as imagens na pasta
    for image_file in image_files:
        # Constrói o caminho completo para a imagem
        image_path = os.path.join(script_dir, image_file)

        # Extrai o texto da imagem
        extracted_text = image_to_text(image_path)

        # Cria um documento do Word e adiciona o texto
        create_word_document(extracted_text, f"output_{os.path.splitext(image_file)[0]}.docx")

        print(f'Texto extraído da imagem {image_file}: {extracted_text}')
        print(f'Documento do Word criado com sucesso para a imagem {image_file}.')
