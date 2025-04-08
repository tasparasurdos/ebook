import pandas as pd
import os
import subprocess
import shutil

# Definir os IDs das planilhas do Google Sheets
google_sheets_tecnologias_id = '1g-zamZFp5FHTGxZOA0vT3ULXjwROhWS0mz0_K1xTvXc'
google_sheets_recursos_id = '1lePYinFlYePVYPwwUqy1Xa1r1FF_l3vaAkRxntSvWq4'

# Construir as URLs de exportação para CSV
url_tecnologias = f'https://docs.google.com/spreadsheets/d/{google_sheets_tecnologias_id}/export?format=csv'
url_recursos = f'https://docs.google.com/spreadsheets/d/{google_sheets_recursos_id}/export?format=csv'

# Verificar e excluir o arquivo ebook.pdf, se existir
if os.path.exists('ebook.pdf'):
    os.remove('ebook.pdf')  # Excluir o arquivo

# Verificar se a pasta .tmp existe e excluí-la
if os.path.exists('.tmp'):
    shutil.rmtree('.tmp')  # Excluir a pasta e seu conteúdo

# Criar a pasta novamente
os.makedirs('.tmp')

# Obter o diretório base do projeto
base_dir = os.path.abspath('.')

# Carregar templates HTML
try:
    template_categoria = open('conteudos/categoria_template.html', 'r', encoding='utf-8').read()
    template_tecnologia = open('conteudos/tecnologia_template.html', 'r', encoding='utf-8').read()
    template_recursos = open('conteudos/recursos.html', 'r', encoding='utf-8').read()
    template_conclusao = open('conteudos/conclusao.html', 'r', encoding='utf-8').read()
    template_contracapa = open('conteudos/contracapa.html', 'r', encoding='utf-8').read()
    print("Templates carregados com sucesso.")
except FileNotFoundError as e:
    print(f"Erro: Arquivo não encontrado - {e}")
    exit(1)

# Ler dados de tecnologias e recursos a partir das URLs do Google Sheets
try:
    tecnologias_df = pd.read_csv(url_tecnologias)
    print("Dados de 'tecnologias' carregados com sucesso do Google Sheets.")
except Exception as e:
    print(f"Erro ao carregar 'tecnologias' do Google Sheets: {e}")
    exit(1)

try:
    recursos_df = pd.read_csv(url_recursos)
    print("Dados de 'recursos' carregados com sucesso do Google Sheets.")
except Exception as e:
    print(f"Erro ao carregar 'recursos' do Google Sheets: {e}")
    exit(1)

# Agrupar tecnologias por categoria
grupos = tecnologias_df.groupby('categoria')

# Função auxiliar para formatar a coluna 'link'
def formatar_links(link_str):
    if not link_str or pd.isna(link_str):
        return ""
    pares = link_str.split(';')
    lista_html = "<ul>"
    for par in pares:
        if par:
            plataforma, link = par.split(',')
            lista_html += f"<li>{plataforma}: <a href='{link}' class='text-decoration-none'>{link}</a></li>"
    lista_html += "</ul>"
    return lista_html

# Iniciar o conteúdo HTML com caminhos absolutos
html_content = f"""
<html>
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" type="text/css" href="file://{os.path.join(base_dir, 'conteudos', 'style.css')}">
    <link rel="stylesheet" type="text/css" href="../conteudos/style.css">
</head>
<body>
"""

# Adicionar capa
with open('conteudos/capa.html', 'r', encoding='utf-8') as f:
    html_content += f"<div style='page-break-after: always;'>{f.read()}</div>"

# Adicionar introdução
with open('conteudos/introducao.html', 'r', encoding='utf-8') as f:
    html_content += f"<div style='page-break-after: always;'>{f.read()}</div>"

# Adicionar instruções
with open('conteudos/instrucoes.html', 'r', encoding='utf-8') as f:
    html_content += f"<div style='page-break-after: always;'>{f.read()}</div>"

# Adicionar categorias e tecnologias
for categoria, grupo in grupos:
    descricao_categ = grupo['categoria_descricao'].iloc[0]
    conteudo_categ = template_categoria.format(categoria=categoria, descricao_categoria=descricao_categ)
    html_content += f"<div style='page-break-after: always;'>{conteudo_categ}</div>"

    for index, row in grupo.sort_values('titulo').iterrows():
        # Criar um dicionário com os dados da linha
        data = row.to_dict()
        # Atualizar o caminho da imagem para absoluto
        imagem_path = f"file://{os.path.join(base_dir, 'imagens', data['imagem'])}"
        data['imagem'] = imagem_path
        # Formatar a coluna 'link'
        data['links_formatados'] = formatar_links(data['link'])
        # Formatar o template com o dicionário atualizado
        conteudo_tec = template_tecnologia.format(**data)
        html_content += f"<div style='page-break-after: always;'>{conteudo_tec}</div>"

# Gerar lista de recursos
if recursos_df.empty:
    lista_recursos = "<p>Nenhum recurso disponível.</p>"
else:
    lista_recursos = ""
    for index, row in recursos_df.iterrows():
        lista_recursos += f"<p><strong>{row['titulo']}</strong><br>"
        lista_recursos += f"<span class='text-muted'>{row['descricao']}</span><br>"
        lista_recursos += f"<a href='{row['link']}' class='text-decoration-none'>{row['link']}</a><hr>"
conteudo_recursos = template_recursos.replace('{lista_recursos}', lista_recursos)
html_content += f"<div style='page-break-after: always;'>{conteudo_recursos}</div>"

# Adicionar conclusão
html_content += f"<div style='page-break-after: always;'>{template_conclusao}</div>"

# Adicionar contracapa
html_content += f"<div style='page-break-after: always;'>{template_contracapa}</div>"

html_content += "</body></html>"

# Salvar o HTML completo em um arquivo temporário
with open('.tmp/ebook.html', 'w', encoding='utf-8') as f:
    f.write(html_content)
print("Arquivo '.tmp/ebook.html' gerado com sucesso.")

# Converter HTML para PDF usando wkhtmltopdf
try:
    subprocess.run(['wkhtmltopdf', '--enable-local-file-access', '.tmp/ebook.html', 'ebook.pdf'], check=True)
    print("PDF 'ebook.pdf' gerado com sucesso usando wkhtmltopdf!")
except FileNotFoundError:
    print("Erro: wkhtmltopdf não encontrado. Certifique-se de que está instalado e no PATH.")
    exit(1)
except subprocess.CalledProcessError as e:
    print(f"Erro ao gerar PDF: {e}")
    exit(1)