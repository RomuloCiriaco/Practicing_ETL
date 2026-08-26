# Exercício 1 — Python + MySQL + DataFrame

Parte do repositório **Practicing_ETL**.

Conectar no MySQL, ler a tabela `ALUNO`, transformar em DataFrame Pandas, exibir e calcular a média das notas.

## Como rodar

Na pasta deste exercício:

```bash
cd exercicios/python_mysql_dataframe
cp .env.example .env          # se ainda não tiver .env
pip install -r ../../requirements.txt
docker compose up -d
jupyter notebook exercicio_aluno.ipynb
```

## Arquivos

| Arquivo | Descrição |
|---------|-----------|
| `exercicio_aluno.ipynb` | Solução do exercício |
| `sql/01_create_aluno.sql` | `CREATE TABLE ALUNO` + dados de exemplo |
| `docker-compose.yml` | MySQL 8 via Docker |
| `.env.example` | Modelo de credenciais |

## Tabela ALUNO

```sql
CREATE TABLE ALUNO (
    idAluno INT PRIMARY KEY AUTO_INCREMENT,
    nome VARCHAR(50),
    curso VARCHAR(30),
    nota DECIMAL(4,2)
);
```

## Observação

O arquivo `.env` não sobe no Git. Use o `.env.example` como referência.
