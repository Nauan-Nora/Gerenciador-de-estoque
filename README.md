# Gerenciador de Estoque

Aplicativo desktop para gerenciamento de produtos e movimentações de estoque,
desenvolvido em Python com interface PySide6 e banco de dados PostgreSQL.

## Funcionalidades

- Cadastro, edição e exclusão de produtos
- Pesquisa por código ou nome
- Bipagem rápida de produto por código de barras; a leitura abre a tela de
  movimentação, mas não altera o estoque sem confirmação
- Movimentações de entrada, venda e perda/uso
- Atualização do custo médio ponderado ao registrar entradas
- Validação de saldo para impedir que saídas ou perdas deixem o estoque negativo
- Histórico geral ou por produto, com as movimentações mais recentes
- Indicadores de quantidade de produtos, estoque baixo (até 5 unidades), valor
  em estoque, potencial de venda, lucro potencial e receita dos últimos 30 dias
- Classificação ABC pela receita de vendas dos últimos 30 dias
- Estimativa de dias de estoque com base nas saídas dos últimos 30 dias
- Temas claro e escuro

## Tecnologias e dependências

- **Python 3.10+**: linguagem e biblioteca padrão (`decimal`, `dataclasses`,
  `datetime`, `pathlib` e outras)
- **PySide6**: interface gráfica Qt
- **psycopg 3** (`psycopg[binary]`): conexão e operações com PostgreSQL
- **python-dotenv**: leitura e gravação da configuração local em `.env`
- **PostgreSQL**: persistência de produtos e movimentações

As dependências Python estão declaradas em `requirements.txt`.

## Instalação e configuração

No PowerShell, na pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Na primeira execução, o aplicativo apresenta a configuração do PostgreSQL.
Informe host, porta, nome do banco, usuário e senha. Também é possível criar
`.env` manualmente a partir de `.env.example`, definindo a variável
`DATABASE_URL`, por exemplo:

```env
DATABASE_URL=postgresql://usuario:senha@localhost:5432/gerenciador_estoque
```

O usuário configurado precisa ter permissão para criar o banco caso ele ainda
não exista e para criar/atualizar as tabelas `products` e `stock_movements`.
As credenciais ficam no arquivo local `.env`, que não deve ser compartilhado
nem enviado ao controle de versão.

## Executar

Com o ambiente virtual ativado:

```powershell
python app\main.py
```

## Testes

Execute os testes automatizados com:

```powershell
python -m unittest discover -s tests
```

## Organização do projeto

```text
.
├── app/
│   ├── main.py                     # Inicialização e janela principal
│   ├── database.py                 # Repositório PostgreSQL e regras de persistência
│   ├── models.py                   # Modelos Product e StockMovement
│   ├── mainwindow.ui                # Arquivo de interface Qt
│   └── ui/
│       ├── components/             # Diálogos e cartões da interface
│       ├── formatting.py           # Formatação de valores
│       └── styles/                 # Temas e estilos
├── tests/
│   └── test_inventory.py           # Testes das movimentações de estoque
├── .env.example                    # Modelo da configuração PostgreSQL
├── requirements.txt                # Dependências Python
└── README.md
```

## Principais módulos e funções

- `main()` em `app/main.py`: inicializa o Qt, carrega ou solicita a configuração
  do banco e abre a janela principal.
- `MainWindow` em `app/main.py`: coordena a interface, busca e exibição dos
  produtos e ações de cadastro, edição, exclusão e movimentação.
- `ProductRepository` em `app/database.py`: concentra o acesso ao PostgreSQL.
  Entre seus métodos estão `initialize()`, `list_products()`,
  `get_product_by_code()`, `get_summary()`, `add_product()`,
  `update_product()`, `register_movement()`, `list_movements()` e
  `delete_product()`.
- `Product` e `StockMovement` em `app/models.py`: representam produtos e
  registros de movimentação; `StockMovementError` identifica erros de regra de
  negócio, como saldo insuficiente.
- `ProductDialog` e `StockMovementDialog` em `app/ui/components/`: coletam e
  validam os dados informados nos formulários.

## Regras e observações

- A quantidade inicial cadastrada cria uma movimentação de entrada no histórico.
- O custo médio é atualizado nas entradas; o cadastro/edição do produto não
  substitui o estoque existente por uma edição manual.
- Vendas gravam o preço de venda vigente no momento da movimentação. Perdas/uso
  reduzem o saldo, mas não contam como receita.
- A classe ABC usa os limites acumulados de 80% (A) e 95% (B) da receita de
  vendas. Produtos sem receita no período ficam na classe C.
- A estimativa de dias de estoque usa a quantidade atual e as saídas (vendas e
  perdas/uso) dos últimos 30 dias. Sem histórico nesse período, é exibida como
  indisponível.

## Uso de inteligência artificial

Ferramentas de inteligência artificial foram utilizadas como apoio durante o
desenvolvimento para analisar e otimizar código e documentação. As sugestões
foram revisadas e validadas; a aplicação não utiliza IA em tempo de execução e
não envia os dados do estoque a serviços de IA.
