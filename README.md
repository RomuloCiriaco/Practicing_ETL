# Practicing_ETL

Projeto de práticas de ETL e integração de dados com Python, Pandas e MySQL.

## Arquitetura

```
Practicing_ETL/
├── notebooks/
│   └── tratamento_bases.ipynb
├── dados/
│   ├── brutos/
│   └── tratados/
├── exercicios/
│   ├── python_mysql_dataframe/   # Python + MySQL + DataFrame
│   ├── imap_migracao/            # Verificação / migração IMAP (genérico)
│   └── inventario_maquinas/      # Consolidação de inventário (genérico)
├── requirements.txt
└── README.md
```

| Pasta | Conteúdo |
|-------|----------|
| `notebooks/` | Notebooks de ETL |
| `dados/` | Bases brutas e tratadas (estudo) |
| `exercicios/` | Atividades complementares |

**Dados sensíveis de trabalho** (senhas, planilhas de cliente, inventários reais)
ficam **fora** deste repositório (ex.: pendrive `EmailSato` / `SATO-INVENTARIO`).

## Instalação

```bash
pip install -r requirements.txt
```

## 1) Tratamento das bases (questionário)

```bash
jupyter notebook notebooks/tratamento_bases.ipynb
```

## 2) Exercício Python + MySQL + DataFrame

```bash
cd exercicios/python_mysql_dataframe
cp .env.example .env
docker compose up -d
jupyter notebook exercicio_aluno.ipynb
```

## 3) IMAP (verificação / migração)

```bash
cd exercicios/imap_migracao
cp .env.example .env
# aponte PLANILHA_* e USB_OUTPUT_DIR para sua pasta de dados
python3 verificar_conta.py --modo unica
```

Detalhes: [`exercicios/imap_migracao/README.md`](exercicios/imap_migracao/README.md).

## 4) Inventário de máquinas

```bash
cd exercicios/inventario_maquinas
# dados reais em dados/ (gitignored) ou no pendrive SATO-INVENTARIO
python3 consolidar_inventario.py
```
