# Projeto: Gerador de Relatório PDF

## Descrição
Este projeto conecta-se a um banco de dados SQL Server, extrai informações sobre produtos e gera um relatório em PDF com essas informações. O relatório inclui ícones que indicam o status do produto e está formatado com tabelas estilizadas.

## Instalação

### Requisitos
- Python 3.x
- Bibliotecas Python: `pandas`, `pyodbc`, `reportlab`

### Passos de Instalação
1. Clone o repositório para o seu ambiente local:
   ```bash
   git clone https://github.com/marcus-21/paginated_report.git
   cd paginated_report
    ```

2. Crie um ambiente virtual e ative-o:
    ```python
      python -m venv venv
      source venv/bin/activate  # Linux/macOS
      .\venv\Scripts\activate  # Windows
    ```

3. Instale as dependências necessárias:
   ```python
    pip install pandas pyodbc reportlab
    ```
4. Certifique-se de ter os ícones necessários (`aceitar.png`, `recarregar.png`, `cancelar.pn`) na pasta `./icon/`.

### Uso
1. Atualize os detalhes da conexão com o banco de dados no script:
    ```python
    server = 'ENDERECO_DO_SERVIDOR'
    database = 'NOME_DO_BANCO_DE_DADOS'
    username = 'SEU_NOME_DE_USUARIO'
    password = 'SUA_SENHA'
    ```

2. Execute o script:
    ```bash
    python relatório.py
    ```
3. O PDF será gerado com o nome `relatorio_personalizado.pdf` no diretório atual.

4.Certifique-se de deletar o `relatorio_personalizado.pdf` criado como exemplo para que não de erro ao seu ser gerado.

### Contribuição
Para contribuir com este projeto, siga os passos abaixo:

1. Faça um fork do projeto.
2. Crie uma branch para sua feature (git checkout -b feature/AmazingFeature).
3. Commit suas mudanças (git commit -m 'Add some AmazingFeature').
4. Faça o push para a branch (git push origin feature/AmazingFeature).
5. Abra um Pull Request.

### Autores
* Marcus Silva - Desenvolvedor Principal - [Meu Perfil](https://github.com/marcus-21)

### Tecnologias Utilizadas
* Python
  * Pandas
  * pyodbc
  * ReportLab
 
### Requisitos
* Sistema Operacional: Windows, macOS, ou Linux
* Banco de Dados: SQL Server

### Resultado 
[Relatório_Personalizado.pdf](https://github.com/marcus-21/paginated_report/blob/main/relatorio_personalizado.pdf)
   
