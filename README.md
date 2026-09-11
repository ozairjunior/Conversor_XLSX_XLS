# Conversor XLSX → XLS

Aplicação desktop para converter planilhas modernas do formato `.xlsx` para o formato legado `.xls` (Excel 97–2003).

A ferramenta possui uma interface gráfica em Tkinter, permite selecionar as pastas de entrada e saída e divide automaticamente planilhas grandes para respeitar o limite do formato `.xls`.

## Funcionalidades

- Conversão em lote de todos os arquivos `.xlsx` de uma pasta.
- Preservação das abas de cada planilha.
- Exportação para o formato `.xls` compatível com Excel 97–2003.
- Divisão automática de abas com mais de 65.535 linhas de dados.
- Configuração do limite de linhas por arquivo.
- Seleção visual das pastas de entrada e saída.
- Barra de progresso e log detalhado da conversão.
- Processamento em segundo plano para manter a interface responsiva.
- Identificação de arquivos convertidos e arquivos com erro.

## Requisitos

- Python 3.9 ou superior
- Tkinter, normalmente incluído na instalação do Python para Windows
- Bibliotecas Python:
  - `pandas`
  - `openpyxl`
  - `xlwt`

## Instalação

Clone o repositório e acesse a pasta do projeto:

```bash
git clone https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
cd SEU_REPOSITORIO
```

Instale as dependências:

```bash
pip install pandas openpyxl xlwt
```

No Windows, caso o comando `pip` não esteja disponível, utilize:

```bash
python -m pip install pandas openpyxl xlwt
```

## Como executar

Execute o arquivo principal:

```bash
python conversor_app.py
```

Na aplicação:

1. Informe ou selecione a pasta que contém os arquivos `.xlsx`.
2. Informe ou selecione a pasta onde os arquivos `.xls` serão salvos.
3. Defina o limite de linhas, se necessário.
4. Mantenha habilitada a opção de divisão automática para separar planilhas grandes.
5. Clique em **Executar conversão**.

## Comportamento da conversão

Quando todas as abas estão dentro do limite configurado, cada arquivo `.xlsx` gera um único arquivo `.xls`, mantendo suas abas.

Quando uma ou mais abas ultrapassam o limite e a divisão automática está habilitada:

- abas menores são exportadas individualmente;
- abas maiores são divididas em partes;
- os arquivos recebem nomes como `arquivo_Aba_parte_1.xls`.

O limite padrão é de 65.535 linhas de dados, conforme a capacidade do formato `.xls`. A linha de cabeçalho é gravada adicionalmente no arquivo exportado.

## Estrutura do projeto

```text
.
├── conversor_app.py
└── README.md
```

## Observações

- Apenas arquivos com extensão `.xlsx` são processados.
- A pasta de saída é criada automaticamente quando não existir.
- Os arquivos originais não são alterados.
- O formato `.xls` é antigo e possui limitações de linhas, colunas e recursos em comparação ao `.xlsx`.
- Fórmulas, estilos avançados, gráficos e outros recursos específicos do Excel podem não ser preservados, pois os dados são lidos e gravados como tabelas.

## Licença

Este projeto ainda não possui uma licença definida. Adicione uma licença ao repositório caso pretenda permitir seu uso, cópia ou modificação por terceiros.
