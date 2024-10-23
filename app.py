import openpyxl
import flet as ft
from barcode import Code128
from barcode.writer import ImageWriter
import io
from PIL import Image
import base64

# Função para carregar a planilha e buscar o funcionário
def buscar_funcionario(matricula, dados_planilha):
    matricula = matricula.lstrip('0')  # Remover zeros iniciais
    
    # Iterar sobre os dados carregados
    for row in dados_planilha:
        nome = row[2]  # Coluna C (Nome do Funcionário)
        chapa = str(row[1]).lstrip('0')  # Coluna B (Matrícula)
        cesta = row[7]  # Coluna H (Cesta)
        motivo = row[8]  # Coluna I (Motivo)
        
        if str(chapa) == matricula:
            if cesta == 'SIM':
                # Gerar código de barras
                return f"Funcionário: {nome}\nTem direito à cesta básica.\nMotivo: {motivo}", "check_circle", ft.colors.GREEN, chapa
            else:
                return f"Funcionário: {nome}\nNão tem direito à cesta básica.\nMotivo: {motivo}", "error", ft.colors.RED, None
    
    return "Funcionário não encontrado.", "error_outline", ft.colors.ORANGE, None

# Função para gerar o código de barras com tamanho ajustado
def gerar_codigo_barras(chapa):
    output = io.BytesIO()
    barcode = Code128(chapa, writer=ImageWriter())
    # Ajustar a largura e altura do código de barras (aumentado conforme solicitado)
    barcode.write(output, {"module_width": 0.8, "module_height": 40, "font_size": 18})
    output.seek(0)
    
    return output

# Função para converter imagem em base64
def imagem_para_base64(imagem):
    return base64.b64encode(imagem.getvalue()).decode('utf-8')

# Função para carregar a planilha
def carregar_planilha(caminho_planilha):
    try:
        wb = openpyxl.load_workbook(caminho_planilha)
        sheet = wb.active
        dados = list(sheet.iter_rows(min_row=2, values_only=True))
        return dados
    except Exception as e:
        return f"Erro ao carregar a planilha: {e}"

# Função principal da interface
def main(page: ft.Page):
    page.title = "Consulta Cesta Básica"
    page.padding = 20
    page.bgcolor = ft.colors.WHITE
    page.scroll = "auto"  # Habilita scroll para telas menores (responsividade)

    caminho_planilha = None
    dados_planilha = None
    img_barcode = ft.Image(width=400, height=120)  # Aumentando o tamanho da imagem

    header = ft.Text("Consulta De Cesta Básica", size=30, weight=ft.FontWeight.BOLD, color=ft.colors.BLUE, text_align="center")

    # Campo de input com Enter
    matricula_input = ft.TextField(
        label="Matrícula do Funcionário", 
        width=300,
        color=ft.colors.BLACK, 
        bgcolor=ft.colors.WHITE, 
        border_color=ft.colors.BLACK,
        on_submit=lambda e: buscar_click(e),  # Captura do Enter
    )

    resultado_icon = ft.Icon(name="info_outline", size=50, visible=True)  # Ícone inicial oculto
    resultado_text = ft.Text(size=18, color=ft.colors.BLACK, visible=False)  # Texto inicial oculto
    matricula_text = ft.Text(size=18, color=ft.colors.BLACK, visible=False)  # Texto para a matrícula

    # Função para carregar a planilha
    def selecionar_planilha(e):
        nonlocal caminho_planilha, dados_planilha
        caminho_planilha = e.files[0].path if e.files else None
        if caminho_planilha:
            resultado_text.value = "Carregando planilha, por favor aguarde..."
            resultado_text.visible = True
            resultado_icon.name = "hourglass_empty"
            resultado_icon.visible = True
            page.update()
            dados_planilha = carregar_planilha(caminho_planilha)
            if isinstance(dados_planilha, list):
                resultado_text.value = "Planilha carregada com sucesso!"
                resultado_icon.name = "check_circle"
                resultado_icon.color = ft.colors.GREEN
            else:
                resultado_text.value = dados_planilha
                resultado_icon.name = "error_outline"
                resultado_icon.color = ft.colors.RED
        else:
            resultado_text.value = "Nenhuma planilha selecionada."
            resultado_icon.name = "error_outline"
            resultado_icon.color = ft.colors.ORANGE
            
        # Definindo a imagem com visibilidade inicialmente oculta
        img_barcode = ft.Image(width=400, height=120, visible=True)  # Aumentando o tamanho da imagem

        page.update()
        

    # Função para buscar o funcionário
    def buscar_click(e):
        if caminho_planilha and dados_planilha and matricula_input.value:
            matricula = matricula_input.value
            resultado, icon_name, icon_color, chapa = buscar_funcionario(matricula, dados_planilha)
            resultado_text.value = resultado
            resultado_text.visible = True  # Torna o texto visível ao buscar
            resultado_icon.name = icon_name
            resultado_icon.color = icon_color
            resultado_icon.visible = True  # Torna o ícone visível ao buscar
            
            if chapa:
                barcode_img = gerar_codigo_barras(chapa)
                img_barcode.src_base64 = imagem_para_base64(barcode_img)  # Exibe o código de barras gerado
                img_barcode.visible = True  # Agora a imagem é visível
                # matricula_text.value = chapa  # Exibe a matrícula abaixo do código de barras
                matricula_text.visible = True
                matricula_input.value = ""  # Limpar o campo após a busca
                
            else:
                img_barcode.visible = False
                matricula_text.visible = False
        else:
            resultado_text.value = "Por favor, selecione a planilha e insira a matrícula."
            resultado_text.visible = True
            resultado_icon.name = "error_outline"
            resultado_icon.color = ft.colors.ORANGE
            resultado_icon.visible = True
            img_barcode.visible = False  # Oculta a imagem se não houver dados válidos
        page.update()

    carregar_button = ft.ElevatedButton("Carregar Planilha", on_click=lambda _: planilha_picker.pick_files(allow_multiple=False))
    buscar_button = ft.ElevatedButton(text="Buscar", on_click=buscar_click)

    planilha_picker = ft.FilePicker(on_result=selecionar_planilha)
    page.overlay.append(planilha_picker)

    # Layout responsivo com rolagem e espaçamento usando ft.Container
    page.add(
        ft.Row([header], alignment=ft.MainAxisAlignment.CENTER),  # Centralizando o título
        ft.Row([carregar_button], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([matricula_input], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([buscar_button], alignment=ft.MainAxisAlignment.CENTER),
        ft.Divider(),
        ft.Row([resultado_icon, resultado_text], alignment=ft.MainAxisAlignment.CENTER),
        ft.Container(height=20),  # Espaço extra
        ft.Row([img_barcode], alignment=ft.MainAxisAlignment.CENTER),  # Exibição do código de barras
        ft.Container(height=10),  # Mais espaço abaixo do código de barras
        ft.Row([matricula_text], alignment=ft.MainAxisAlignment.CENTER),  # Exibição da matrícula
        ft.Container(height=30)  # Espaço abaixo da matrícula
    )

# Inicializa a aplicação
ft.app(target=main)
