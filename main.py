from pathlib import Path
from shutil import copyfile
from datetime import datetime
import win32com.client
import time
import tempfile
from PIL import ImageGrab
import sys
import win32clipboard
import re

BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

WD_EXPORT_FORMAT_PDF = 17


# Procura os modelos "Modelo <Obra>.docx" disponíveis na pasta Modelo/
def listar_modelos():
    pasta_modelo = BASE / "Modelo"
    return sorted(pasta_modelo.glob("Modelo *.docx"))


def extrair_nome_obra(caminho_modelo):
    return caminho_modelo.stem.removeprefix("Modelo ").strip()


def escolher_modelo():
    modelos = listar_modelos()
    if not modelos:
        raise FileNotFoundError(f"Nenhum modelo encontrado em {BASE / 'Modelo'}")

    if len(modelos) == 1:
        return modelos[0]

    print("Obras disponíveis:")
    for i, modelo in enumerate(modelos, start=1):
        print(f"  {i}. {extrair_nome_obra(modelo)}")

    while True:
        escolha = input("Escolha o número da obra: ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(modelos):
            return modelos[int(escolha) - 1]
        print("Opção inválida.")


# Pega o Template da obra escolhida e cria uma cópia
def criar_relatorio(modelo):
    nome_obra = extrair_nome_obra(modelo)

    pasta_obra = BASE / "Relatórios" / nome_obra
    pasta_obra.mkdir(parents=True, exist_ok=True)

    agora = datetime.now()
    nome = agora.strftime(f"Relatório {nome_obra} %d-%m-%Y.docx")
    destino = pasta_obra / nome

    copyfile(modelo, destino)

    print("Relatório criado:")
    print(destino)

    return destino

# Lê no próprio documento quais números de bookmark "<prefixo>N" existem
def obter_numeros_bookmark(documento, prefixo):
    numeros = []
    for i in range(1, documento.Bookmarks.Count + 1):
        correspondencia = re.fullmatch(rf"{prefixo}(\d+)", documento.Bookmarks(i).Name)
        if correspondencia:
            numeros.append(int(correspondencia.group(1)))
    return sorted(numeros)


# Abre relatótio no word e adiciona as datas automaticamente
def abrir_relatorio_no_word(caminho):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True

    documento = word.Documents.Open(str(caminho.resolve()))
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    for numero in obter_numeros_bookmark(documento, "Data"):
        documento.Bookmarks(f"Data{numero}").Range.Text = agora

    documento.Save()

    return word, documento


def limpar_clipboard():
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
    finally:
        win32clipboard.CloseClipboard()


def inserir_foto(documento, numero, caminho_temp):
    marcador = documento.Bookmarks(f"Foto{numero}").Range

    if marcador.InlineShapes.Count > 0:
        marcador.InlineShapes(1).Delete()

    marcador.InlineShapes.AddPicture(
        caminho_temp,
        LinkToFile=False,
        SaveWithDocument=True
    )

    documento.Bookmarks.Add(f"Foto{numero}", marcador)


def exportar_pdf(documento, caminho_docx):
    caminho_pdf = caminho_docx.with_suffix(".pdf")
    documento.ExportAsFixedFormat(str(caminho_pdf.resolve()), WD_EXPORT_FORMAT_PDF)
    print("PDF exportado:")
    print(caminho_pdf)


def main():
    modelo = escolher_modelo()
    arquivo = criar_relatorio(modelo)
    word = None
    documento = None

    try:
        word, documento = abrir_relatorio_no_word(arquivo)

        fotos_do_modelo = obter_numeros_bookmark(documento, "Foto")
        if not fotos_do_modelo:
            raise RuntimeError("Nenhum marcador 'FotoN' encontrado no modelo.")

        print(f"Aguardando capturas... ({len(fotos_do_modelo)} fotos no modelo)")

        for numero in fotos_do_modelo:
            print(f"Aguardando foto {numero}... (copie a imagem para a área de transferência)")

            imagem = None
            while imagem is None:
                imagem = ImageGrab.grabclipboard()
                if imagem is None:
                    time.sleep(0.5)

            caminho_temp = tempfile.gettempdir() + f"\\foto{numero}.png"
            imagem.save(caminho_temp)

            inserir_foto(documento, numero, caminho_temp)

            print(f"[{datetime.now().strftime('%H:%M:%S')}] Foto {numero} inserida!")

            limpar_clipboard()

        documento.Save()
        exportar_pdf(documento, arquivo)

        print("Relatório finalizado!")

    except Exception as erro:
        print(f"Erro ao gerar o relatório: {erro}")
        raise

    finally:
        if documento is not None:
            try:
                documento.Save()
            except Exception:
                pass
            documento.Close(SaveChanges=False)
        if word is not None:
            word.Quit()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        pass
    input("Pressione Enter para sair...")
  