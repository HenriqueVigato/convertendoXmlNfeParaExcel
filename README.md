📄 Lançador de Notas Fiscais XML
Ferramenta desktop para importação automática de Notas Fiscais Eletrônicas (NF-e) em formato XML, com gravação dos dados em uma planilha Excel compartilhada em rede.

🚀 Funcionalidades

Seleção de um ou mais arquivos XML de NF-e via explorador de arquivos do Windows
Validação da estrutura do XML antes de processar (rejeita arquivos que não sejam NF-e)
Extração automática dos dados da nota:

Nome do fornecedor
Número da NF-e
Data de emissão
Valor total
Responsável pelo frete (emitente ou destinatário)
Boletos/duplicatas com datas de vencimento e valores


Confirmação interativa dos dados extraídos antes de gravar
Gravação dos dados em planilha Excel em diretório de rede compartilhado
Abertura automática da planilha ao final do processo


🗂️ Estrutura do Fluxo
Usuário seleciona XML(s)
        ↓
Validação da estrutura NF-e
        ↓
Leitura e extração dos dados via xmltodict
        ↓
Exibição dos dados para conferência pelo usuário
        ↓
  Confirmação (s/n)
        ↓
  Gravação no Excel
        ↓
Pergunta se há mais notas a importar
        ↓
  Abre o Excel para conferência final

📋 Pré-requisitos

Python 3.x
Bibliotecas necessárias:

bashpip install xmltodict openpyxl

Microsoft Excel instalado (para abertura automática da planilha)
Acesso à unidade de rede Z:\ com a seguinte estrutura:

Z:\Compartilhado\
├── 1-ARQUIVO XML NOTA ENTRADA\   ← XMLs das notas fiscais
└── Notas de entrada\
    └── Notas de entrada automatico.xlsx  ← Planilha de destino

⚙️ Configuração
Os caminhos de rede ficam em um arquivo separado chamado config.json, que deve estar na mesma pasta do script:
json{
    "caminho_excel": "Z:\\Compartilhado\\Notas de entrada\\Notas de entrada automatico.xlsx",
    "caminho_xml": "Z:\\Compartilhado\\1-ARQUIVO XML NOTA ENTRADA"
}

⚠️ Nunca versione o config.json em repositórios públicos. O arquivo .gitignore já está configurado para bloqueá-lo automaticamente.

Se os caminhos mudarem, edite apenas o config.json — o código não precisa ser alterado.

▶️ Como usar

Certifique-se de que o arquivo config.json está na mesma pasta que o script.
Execute o script:

bash   python lancador_xml.py

O explorador de arquivos será aberto — selecione um ou mais arquivos .xml de NF-e.
Confira os dados exibidos no terminal para cada nota e confirme com s ou cancele com n.
Os dados confirmados serão gravados automaticamente na planilha Excel.
Ao final, o programa pergunta se há mais notas a importar.
Após encerrar, a planilha Excel será aberta automaticamente para conferência.


📊 Dados gravados na planilha
ColunaConteúdoFornecedorNome + data de lançamentoNúmero NF-eNúmero da nota fiscalData de emissãoData no formato DD/MM/AAAAValorValor total da nota (numérico)Freteemitente ou destinatarioBoleto(s)Data de vencimento e valor de cada duplicata

🔒 Segurança

Caminhos de rede isolados no config.json, fora do código-fonte
O .gitignore impede que o config.json seja enviado acidentalmente para repositórios
XMLs são validados antes do processamento — arquivos que não sejam NF-e são rejeitados
Erros tratados com tipos específicos (KeyError, OSError, ValueError), evitando falhas silenciosas


⚠️ Observações

O script foi desenvolvido para ambiente Windows com rede mapeada na unidade Z:\.
Caso não existam duplicatas registradas na NF-e, a coluna de boletos exibirá uma mensagem orientando a verificação manual.
O valor monetário é formatado no padrão brasileiro (ex: 1.250,00).
Arquivos XML devem estar no padrão NF-e (nfeProc).
