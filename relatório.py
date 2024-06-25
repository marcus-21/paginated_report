import os
import pandas as pd
import pyodbc
from reportlab.lib.pagesizes import landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# Função para adicionar ícones baseados no status
def status_icon(status):
    icons_dir = './icon/'
    if status == 'Concluído':
        icon_path = os.path.join(icons_dir, 'aceitar.png')
    elif status == 'Em andamento':
        icon_path = os.path.join(icons_dir, 'recarregar.png')
    elif status == 'Atrasado' or status == 'Início Atrasado':
        icon_path = os.path.join(icons_dir, 'cancelar.png')
    else:
        return ''
    
    if os.path.exists(icon_path):
        return Image(icon_path, width=15, height=15)
    else:
        return None

# Função para tratar valor nulo nos comentários
def format_comments(comments):
    if comments:
        return comments.replace('\n', '<br/>')
    return ''

# Detalhes da conexão
server = '**.*.**.**,****'  # substitua pelo nome ou IP do seu servidor com a porta
database = 'loja'  # substitua pelo nome do seu banco de dados
username = '******'  # substitua pelo seu nome de usuário
password = '******'  # substitua pela sua senha

# String de conexão
connection_string = f"""
DRIVER={{ODBC Driver 17 for SQL Server}};
SERVER={server};
DATABASE={database};
UID={username};
PWD={password};
TrustServerCertificate=yes;
"""

# Conectar ao banco de dados e criar o PDF
try:
    conexao = pyodbc.connect(connection_string)
    cursor = conexao.cursor()

    # Executar a consulta SQL
    cursor.execute("""
        SELECT 
        CONCAT(estrutura, ' - ', descricao) AS 'Produto',
        tipo as 'Tipo',
        CASE 
            WHEN p.terminoEfetivo IS NOT NULL AND p.inicioEfetivo IS NOT NULL THEN 'Concluído'
            WHEN p.inicioEfetivo IS NOT NULL AND p.terminoEfetivo IS NULL THEN 'Em andamento'
            WHEN GETDATE() > p.inicioPrevisto THEN 'Atrasado'
            WHEN p.inicioEfetivo IS NULL AND GETDATE() > p.inicioPrevisto AND p.terminoEfetivo IS NULL THEN 'Início Atrasado'
        END AS Status,
        (SELECT nome FROM responsavel r WHERE r.id = p.responsavelId) AS 'Responsável',
        FORMAT(inicioPrevisto, 'dd/MM/yyyy') AS 'Início Previsto',
        FORMAT(terminoPrevisto, 'dd/MM/yyyy') AS 'Término Previsto',
        FORMAT(inicioEfetivo, 'dd/MM/yyyy') AS 'Início Efetivo',
        FORMAT(terminoEfetivo, 'dd/MM/yyyy') AS 'Término Efetivo',
        FORMAT(updatedAt, 'dd/MM/yyyy') AS 'Última Atualização',
        (SELECT STRING_AGG(CONCAT('Data: ', FORMAT(createdAt, 'dd/MM/yyyy'), CHAR(10), 'Comentário: ', c.texto), CHAR(10))
        FROM comentario_produtos c WHERE c.produtosId = p.id) AS ultimosComentarios
        FROM produtos p
        WHERE  LEN(estrutura) > 2 AND cancelada IS NULL
        ORDER BY estrutura;
    """)

    # Obter os resultados da consulta e criar um DataFrame
    rows = cursor.fetchall()
    columns = [column[0] for column in cursor.description]
    df = pd.DataFrame.from_records(rows, columns=columns)

    # Fechar a conexão com o banco de dados
    cursor.close()
    conexao.close()

    # Adicionar coluna com ícones
    df['Ícone'] = df['Status'].apply(status_icon)

    # Tamanho personalizado do PDF
    custom_width = 1200  # largura personalizada em pontos
    custom_height = 800  # altura personalizada em pontos

    # Estilos de parágrafos
    styles = getSampleStyleSheet()
    style_table_header = ParagraphStyle(name='TableHeader', textColor='white', parent=styles['Normal'], valign='middle', alignment=TA_CENTER, fontSize=11)
    style_table_data = ParagraphStyle(name='TableData', parent=styles['Normal'], alignment=TA_CENTER, valign='middle', fontSize=10)

    # Função para criar o PDF com tamanho personalizado
    def create_pdf(df, filename):
        doc = SimpleDocTemplate(filename, pagesize=(custom_width, custom_height))
        elements = []

        # Adicionar título do relatório com imagem
        title_style = ParagraphStyle(name='Title', parent=styles['Title'], alignment=TA_CENTER, fontSize=16)
        title = Paragraph("Relatório", title_style)
        img = Image('./icon/logo.jpg', width=50, height=50)  # ajustei o tamanho da imagem
        
        # Tabela com imagem e título centralizados
        title_table = Table([[img, title]], colWidths=[30, 100])
        title_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))

        elements.append(title_table)
        
        # Adicionar espaço
        elements.append(Spacer(1, 15))

        # Cabeçalho da tabela
        table_data = []
        table_data.append([Paragraph('<b>Ícone</b>', style_table_header),
                           Paragraph('<b>Produto</b>', style_table_header),
                           Paragraph('<b>Tipo</b>', style_table_header),
                           Paragraph('<b>Status</b>', style_table_header),
                           Paragraph('<b>Responsável</b>', style_table_header),
                           Paragraph('<b>Início Previsto</b>', style_table_header),
                           Paragraph('<b>Término Previsto</b>', style_table_header),
                           Paragraph('<b>Início Efetivo</b>', style_table_header),
                           Paragraph('<b>Término Efetivo</b>', style_table_header),
                           Paragraph('<b>Última Atualização</b>', style_table_header),
                           Paragraph('<b>Últimos Comentários</b>', style_table_header)])

        for _, row in df.iterrows():
            icon = row['Ícone'] if not pd.isnull(row['Ícone']) else ''
            table_data.append([
                icon,
                Paragraph(format_comments(row['Produto']), style_table_data),
                Paragraph(format_comments(row['Tipo']), style_table_data),
                Paragraph(format_comments(row['Status']), style_table_data),
                Paragraph(format_comments(row['Responsável']), style_table_data),
                Paragraph(format_comments(row['Início Previsto']), style_table_data),
                Paragraph(format_comments(row['Término Previsto']), style_table_data),
                Paragraph(format_comments(row['Início Efetivo']), style_table_data),
                Paragraph(format_comments(row['Término Efetivo']), style_table_data),
                Paragraph(format_comments(row['Última Atualização']), style_table_data),
                Paragraph(format_comments(row['ultimosComentarios']), style_table_data),
            ])

        # Tamanho das colunas
        col_widths = [42, 100, 70, 80, 80, 80, 80, 80, 80, 80, 150]

        # Tabela
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ]))
        elements.append(table)

        # Adicionar elementos ao PDF
        doc.build(elements)
        print(f"Relatório exportado como {filename}")

    # Exportar o PDF
    pdf_file = "relatorio_personalizado.pdf"
    create_pdf(df, pdf_file)

except pyodbc.Error as ex:
    print("Erro na conexão com o banco de dados:")
    print(ex)
except Exception as e:
    print("Erro ao gerar o PDF:")
    print(e)
