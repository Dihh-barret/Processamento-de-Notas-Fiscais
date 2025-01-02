# Processamento de Notas Fiscais (PDF) - Extrator de Dados

Este projeto realiza a extração de dados específicos de arquivos PDF de Notas Fiscais, como o valor do serviço, o nome do tomador do serviço e outros detalhes, e gera um relatório em formato Excel com esses dados. Ele é útil para quem precisa organizar e processar grandes volumes de NFs emitidas em PDF. 

## Funcionalidades

- **Extração do valor do serviço**: O valor do serviço é extraído do PDF, considerando possíveis variações de formato (R$, R$ 10,00, etc.).
- **Extração do nome do tomador do serviço**: O nome do tomador é identificado no campo "Nome / Nome Empresarial", até o marcador de e-mail.
- **Geração de relatório**: O programa gera um arquivo Excel contendo todos os dados extraídos, com o nome do arquivo, caminho e valor do serviço.
- **Processamento em massa**: A aplicação percorre todas as pastas de uma pasta raiz especificada, processando todos os arquivos PDF encontrados.

## Tecnologias Utilizadas

- **Python 3.x**: A linguagem utilizada para o desenvolvimento do script.
- **PyPDF2**: Biblioteca para leitura de arquivos PDF.
- **Pandas**: Para manipulação de dados e criação do arquivo Excel.
- **Regex (expressões regulares)**: Para a extração de dados específicos de dentro do conteúdo dos PDFs.

## Pré-Requisitos

Antes de rodar o projeto, você precisa ter o Python 3.x instalado em sua máquina. Também é necessário instalar as bibliotecas que o projeto utiliza.

### Instalar dependências

Crie um ambiente virtual (opcional, mas recomendado) e instale as dependências necessárias com o seguinte comando:

```bash
pip install -r requirements.txt
```
## Como Usar

1. **Baixe ou clone este repositório**: Você pode baixar ou clonar o repositório para sua máquina.
    
``` bash
    git clone https://github.com/seu-usuario/nome-do-repositorio.git cd nome-do-repositorio
```
2. **Prepare sua pasta de Notas Fiscais**: Coloque seus arquivos PDF de Notas Fiscais em uma pasta local, por exemplo, `NFs emitidas`.
    
3. **Configure o caminho da pasta no código**: No script, altere o valor da variável `pasta_raiz` para o caminho onde seus PDFs estão localizados.
    ```python
    `git clone https://github.com/seu-usuario/nome-do-repositorio.git cd nome-do-repositorio`
    ```

4. **Execute o script**: Para rodar o script e processar os PDFs, execute o seguinte comando no terminal:
	```python
	  python nome_do_script.py
	```
	
5. **Verifique os resultados**: O script irá gerar um arquivo Excel chamado `resultado_notas_fiscais.xlsx` contendo as informações extraídas dos PDFs. O arquivo estará na mesma pasta onde o script foi executado.
    

## Detalhes Técnicos

### Extração de Dados

- **Valor do Serviço**: O valor do serviço é extraído utilizando uma expressão regular que captura os valores numéricos associados ao texto "Valor do Serviço".
    
- **Nome do Tomador**: O nome do tomador é extraído a partir do campo "Nome / Nome Empresarial", levando em conta a possibilidade de ele ser seguido por um e-mail. O regex é ajustado para pegar o nome até esse ponto.
    
- **Relatório Excel**: O relatório gerado contém os seguintes campos:
    
    - **arquivo**: Nome do arquivo PDF.
    - **caminho**: Caminho completo do arquivo PDF.
    - **valor**: Valor extraído do campo "Valor do Serviço".
    - **tomador**: Nome do tomador extraído do campo "Nome / Nome Empresarial".

### Personalização

Se desejar personalizar o projeto para capturar outros dados dos PDFs ou fazer ajustes no formato do relatório, você pode modificar as funções de extração e processamento no código.

## Contribuições

Se você encontrar algum bug ou desejar adicionar novas funcionalidades, sinta-se à vontade para abrir uma **issue** ou **pull request**. Toda contribuição é bem-vinda!
