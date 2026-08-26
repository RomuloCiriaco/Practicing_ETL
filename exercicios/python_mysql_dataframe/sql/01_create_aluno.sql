-- Exercício 1 - Python + MySQL + DataFrame
-- Cria a tabela ALUNO e popula com dados de exemplo

CREATE TABLE IF NOT EXISTS ALUNO (
    idAluno INT PRIMARY KEY AUTO_INCREMENT,
    nome VARCHAR(50),
    curso VARCHAR(30),
    nota DECIMAL(4,2)
);

INSERT INTO ALUNO (nome, curso, nota) VALUES
('Ana Souza', 'ADS', 8.50),
('Bruno Lima', 'ADS', 7.00),
('Carla Mendes', 'SI', 9.25),
('Diego Alves', 'ADS', 6.75),
('Elena Rocha', 'SI', 8.00),
('Felipe Nunes', 'ADS', 5.50),
('Gabriela Costa', 'SI', 9.80),
('Henrique Dias', 'ADS', 7.40);
