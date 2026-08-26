# Exercício — Verificação / migração IMAP

Parte do repositório **Practicing_ETL**.

Lê a estrutura IMAP (origem × destino), gera Excel e permite migração
testada da INBOX. **Dados reais (senhas/planilhas) ficam fora do Git**
(ex.: pasta `EmailSato` no pendrive).

## Como rodar

```bash
cd exercicios/imap_migracao
cp .env.example .env
# edite .env com hosts, paths das planilhas e (modo unica) credenciais
pip install -r ../../requirements.txt
```

### 1 conta (`.env`)

```bash
python3 verificar_conta.py --modo unica
```

### Lote (planilha)

Paths padrão vêm do `.env` (`PLANILHA_ALINHAMENTO`, etc.).

```bash
# teste com 3 contas
python3 verificar_conta.py --modo lote --limite 3

# só INBOX
python3 verificar_conta.py --modo lote --compare-mode INBOX --pular-migrados

# uma conta específica
python3 verificar_conta.py --modo lote --email usuario@exemplo.com.br
```

## Saída

Em `saidas/` (e opcionalmente `USB_OUTPUT_DIR` do `.env`):

| Aba | Conteúdo |
|-----|----------|
| `resumo` | 1 linha por conta |
| `mapeamento` | pastas origem → destino |
| `estrutura_locaweb` / `estrutura_kinghost` | árvores |

## Migração TESTE (só INBOX)

```bash
# simular
python3 migrar_inbox.py --contas /caminho/contas_teste.csv --limite 5 --dry-run

# copiar N msgs
python3 migrar_inbox.py --contas /caminho/contas_teste.csv \
  --email usuario@exemplo.com.br --limite 10
```

Use `contas_teste.csv.example` como modelo. O arquivo real não sobe no Git.

## Arquivos

| Arquivo | Descrição |
|---------|-----------|
| `verificar_conta.py` | CLI unica / lote |
| `migrar_inbox.py` | Migração TESTE da INBOX |
| `imap_estrutura.py` | Núcleo IMAP |
| `planilha_contas.py` | Leitura de planilhas |
| `.env.example` | Modelo de credenciais/caminhos |

## Observação

O `.env` e planilhas com senha **não** entram no Git. Mantenha-os no pendrive
ou em pasta local fora do repositório.
