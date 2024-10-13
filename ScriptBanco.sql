CREATE DATABASE dbCorretagem;
USE dbCorretagem;

-- Tabela de Unidades Federativas (UF) deve ser criada primeiro
CREATE TABLE UF (
    id INT PRIMARY KEY IDENTITY(1,1),
    nome VARCHAR(100) NOT NULL
);

-- Tabela de Cidades deve ser criada em seguida
CREATE TABLE Cidade (
    id INT PRIMARY KEY IDENTITY(1,1),
    nome VARCHAR(100) NOT NULL,
    uf_id INT NOT NULL,
    FOREIGN KEY (uf_id) REFERENCES UF(id)
);

-- Agora podemos criar a Tabela de Clientes
CREATE TABLE Cliente (
    id INT PRIMARY KEY IDENTITY(1,1),
    nome VARCHAR(255) NOT NULL,
    cpf VARCHAR(14) NOT NULL UNIQUE,
    endereco VARCHAR(255) NOT NULL,
    uf_id INT NOT NULL,
    cidade_id INT NOT NULL,
    ativo VARCHAR(1) NOT NULL,
    FOREIGN KEY (uf_id) REFERENCES UF(id),
    FOREIGN KEY (cidade_id) REFERENCES Cidade(id)
);

-- Tabela de Corretores
CREATE TABLE Corretor (
    id INT PRIMARY KEY IDENTITY(1,1),
    nome VARCHAR(255) NOT NULL,
    cpf VARCHAR(14) NOT NULL UNIQUE
);

-- Tabela de Relação entre Corretor e Cliente
CREATE TABLE ClienteCorretor (
    id INT PRIMARY KEY IDENTITY(1,1),
    corretor_id INT NOT NULL,
    cliente_id INT NOT NULL,
    FOREIGN KEY (corretor_id) REFERENCES Corretor(id),
    FOREIGN KEY (cliente_id) REFERENCES Cliente(id)
);


INSERT INTO UF (nome) VALUES
('Acre'),
('Alagoas'),
('Amapá'),
('Amazonas'),
('Bahia'),
('Ceará'),
('Distrito Federal'),
('Espírito Santo'),
('Goiás'),
('Maranhão'),
('Mato Grosso'),
('Mato Grosso do Sul'),
('Minas Gerais'),
('Pará'),
('Paraíba'),
('Paraná'),
('Pernambuco'),
('Piauí'),
('Rio de Janeiro'),
('Rio Grande do Norte'),
('Rio Grande do Sul'),
('Rondônia'),
('Roraima'),
('Santa Catarina'),
('São Paulo'),
('Sergipe'),
('Tocantins');

INSERT INTO Cidade (nome, uf_id) VALUES
('Rio Branco', (SELECT id FROM UF WHERE nome = 'Acre')),
('Maceió', (SELECT id FROM UF WHERE nome = 'Alagoas')),
('Macapá', (SELECT id FROM UF WHERE nome = 'Amapá')),
('Manaus', (SELECT id FROM UF WHERE nome = 'Amazonas')),
('Salvador', (SELECT id FROM UF WHERE nome = 'Bahia')),
('Fortaleza', (SELECT id FROM UF WHERE nome = 'Ceará')),
('Brasília', (SELECT id FROM UF WHERE nome = 'Distrito Federal')),
('Vitória', (SELECT id FROM UF WHERE nome = 'Espírito Santo')),
('Goiânia', (SELECT id FROM UF WHERE nome = 'Goiás')),
('São Luís', (SELECT id FROM UF WHERE nome = 'Maranhão')),
('Cuiabá', (SELECT id FROM UF WHERE nome = 'Mato Grosso')),
('Campo Grande', (SELECT id FROM UF WHERE nome = 'Mato Grosso do Sul')),
('Belo Horizonte', (SELECT id FROM UF WHERE nome = 'Minas Gerais')),
('Belém', (SELECT id FROM UF WHERE nome = 'Pará')),
('João Pessoa', (SELECT id FROM UF WHERE nome = 'Paraíba')),
('Curitiba', (SELECT id FROM UF WHERE nome = 'Paraná')),
('Recife', (SELECT id FROM UF WHERE nome = 'Pernambuco')),
('Teresina', (SELECT id FROM UF WHERE nome = 'Piauí')),
('Rio de Janeiro', (SELECT id FROM UF WHERE nome = 'Rio de Janeiro')),
('Natal', (SELECT id FROM UF WHERE nome = 'Rio Grande do Norte')),
('Porto Alegre', (SELECT id FROM UF WHERE nome = 'Rio Grande do Sul')),
('Porto Velho', (SELECT id FROM UF WHERE nome = 'Rondônia')),
('Boa Vista', (SELECT id FROM UF WHERE nome = 'Roraima')),
('Florianópolis', (SELECT id FROM UF WHERE nome = 'Santa Catarina')),
('São Paulo', (SELECT id FROM UF WHERE nome = 'São Paulo')),
('Aracaju', (SELECT id FROM UF WHERE nome = 'Sergipe')),
('Palmas', (SELECT id FROM UF WHERE nome = 'Tocantins'));
INSERT INTO Cidade (nome, uf_id) VALUES
-- Acre
('Rio Branco', (SELECT id FROM UF WHERE nome = 'Acre')),
('Cruzeiro do Sul', (SELECT id FROM UF WHERE nome = 'Acre')),
('Sena Madureira', (SELECT id FROM UF WHERE nome = 'Acre')),
('Tarauacá', (SELECT id FROM UF WHERE nome = 'Acre')),

-- Alagoas
('Maceió', (SELECT id FROM UF WHERE nome = 'Alagoas')),
('Arapiraca', (SELECT id FROM UF WHERE nome = 'Alagoas')),
('Palmeira dos Índios', (SELECT id FROM UF WHERE nome = 'Alagoas')),
('Penedo', (SELECT id FROM UF WHERE nome = 'Alagoas')),

-- Amapá
('Macapá', (SELECT id FROM UF WHERE nome = 'Amapá')),
('Santana', (SELECT id FROM UF WHERE nome = 'Amapá')),
('Laranjal do Jari', (SELECT id FROM UF WHERE nome = 'Amapá')),
('Tartarugalzinho', (SELECT id FROM UF WHERE nome = 'Amapá')),

-- Amazonas
('Manaus', (SELECT id FROM UF WHERE nome = 'Amazonas')),
('Itacoatiara', (SELECT id FROM UF WHERE nome = 'Amazonas')),
('Parintins', (SELECT id FROM UF WHERE nome = 'Amazonas')),
('Tabatinga', (SELECT id FROM UF WHERE nome = 'Amazonas')),

-- Bahia
('Salvador', (SELECT id FROM UF WHERE nome = 'Bahia')),
('Feira de Santana', (SELECT id FROM UF WHERE nome = 'Bahia')),
('Vitória da Conquista', (SELECT id FROM UF WHERE nome = 'Bahia')),
('Ilhéus', (SELECT id FROM UF WHERE nome = 'Bahia')),

-- Ceará
('Fortaleza', (SELECT id FROM UF WHERE nome = 'Ceará')),
('Caucaia', (SELECT id FROM UF WHERE nome = 'Ceará')),
('Juazeiro do Norte', (SELECT id FROM UF WHERE nome = 'Ceará')),
('Crato', (SELECT id FROM UF WHERE nome = 'Ceará')),

-- Distrito Federal
('Brasília', (SELECT id FROM UF WHERE nome = 'Distrito Federal')),
('Taguatinga', (SELECT id FROM UF WHERE nome = 'Distrito Federal')),
('Ceilândia', (SELECT id FROM UF WHERE nome = 'Distrito Federal')),
('Gama', (SELECT id FROM UF WHERE nome = 'Distrito Federal')),

-- Espírito Santo
('Vitória', (SELECT id FROM UF WHERE nome = 'Espírito Santo')),
('Vila Velha', (SELECT id FROM UF WHERE nome = 'Espírito Santo')),
('Cariacica', (SELECT id FROM UF WHERE nome = 'Espírito Santo')),
('Serra', (SELECT id FROM UF WHERE nome = 'Espírito Santo')),

-- Goiás
('Goiânia', (SELECT id FROM UF WHERE nome = 'Goiás')),
('Aparecida de Goiânia', (SELECT id FROM UF WHERE nome = 'Goiás')),
('Anápolis', (SELECT id FROM UF WHERE nome = 'Goiás')),
('Rio Verde', (SELECT id FROM UF WHERE nome = 'Goiás')),

-- Maranhão
('São Luís', (SELECT id FROM UF WHERE nome = 'Maranhão')),
('Imperatriz', (SELECT id FROM UF WHERE nome = 'Maranhão')),
('São José de Ribamar', (SELECT id FROM UF WHERE nome = 'Maranhão')),
('Caxias', (SELECT id FROM UF WHERE nome = 'Maranhão')),

-- Mato Grosso
('Cuiabá', (SELECT id FROM UF WHERE nome = 'Mato Grosso')),
('Várzea Grande', (SELECT id FROM UF WHERE nome = 'Mato Grosso')),
('Rondonópolis', (SELECT id FROM UF WHERE nome = 'Mato Grosso')),
('Sinop', (SELECT id FROM UF WHERE nome = 'Mato Grosso')),

-- Mato Grosso do Sul
('Campo Grande', (SELECT id FROM UF WHERE nome = 'Mato Grosso do Sul')),
('Dourados', (SELECT id FROM UF WHERE nome = 'Mato Grosso do Sul')),
('Três Lagoas', (SELECT id FROM UF WHERE nome = 'Mato Grosso do Sul')),
('Corumbá', (SELECT id FROM UF WHERE nome = 'Mato Grosso do Sul')),

-- Minas Gerais
('Belo Horizonte', (SELECT id FROM UF WHERE nome = 'Minas Gerais')),
('Uberlândia', (SELECT id FROM UF WHERE nome = 'Minas Gerais')),
('Contagem', (SELECT id FROM UF WHERE nome = 'Minas Gerais')),
('Juiz de Fora', (SELECT id FROM UF WHERE nome = 'Minas Gerais')),

-- Pará
('Belém', (SELECT id FROM UF WHERE nome = 'Pará')),
('Ananindeua', (SELECT id FROM UF WHERE nome = 'Pará')),
('Marabá', (SELECT id FROM UF WHERE nome = 'Pará')),
('Santana do Araguaia', (SELECT id FROM UF WHERE nome = 'Pará')),

-- Paraíba
('João Pessoa', (SELECT id FROM UF WHERE nome = 'Paraíba')),
('Campina Grande', (SELECT id FROM UF WHERE nome = 'Paraíba')),
('Patos', (SELECT id FROM UF WHERE nome = 'Paraíba')),
('Santa Rita', (SELECT id FROM UF WHERE nome = 'Paraíba')),

-- Paraná
('Curitiba', (SELECT id FROM UF WHERE nome = 'Paraná')),
('Londrina', (SELECT id FROM UF WHERE nome = 'Paraná')),
('Maringá', (SELECT id FROM UF WHERE nome = 'Paraná')),
('Ponta Grossa', (SELECT id FROM UF WHERE nome = 'Paraná')),

-- Pernambuco
('Recife', (SELECT id FROM UF WHERE nome = 'Pernambuco')),
('Olinda', (SELECT id FROM UF WHERE nome = 'Pernambuco')),
('Caruaru', (SELECT id FROM UF WHERE nome = 'Pernambuco')),
('Jaboatão dos Guararapes', (SELECT id FROM UF WHERE nome = 'Pernambuco')),

-- Piauí
('Teresina', (SELECT id FROM UF WHERE nome = 'Piauí')),
('Parnaíba', (SELECT id FROM UF WHERE nome = 'Piauí')),
('Picos', (SELECT id FROM UF WHERE nome = 'Piauí')),
('Floriano', (SELECT id FROM UF WHERE nome = 'Piauí')),

-- Rio de Janeiro
('Rio de Janeiro', (SELECT id FROM UF WHERE nome = 'Rio de Janeiro')),
('Niterói', (SELECT id FROM UF WHERE nome = 'Rio de Janeiro')),
('Duque de Caxias', (SELECT id FROM UF WHERE nome = 'Rio de Janeiro')),
('Nova Iguaçu', (SELECT id FROM UF WHERE nome = 'Rio de Janeiro')),

-- Rio Grande do Norte
('Natal', (SELECT id FROM UF WHERE nome = 'Rio Grande do Norte')),
('Mossoró', (SELECT id FROM UF WHERE nome = 'Rio Grande do Norte')),
('Parnamirim', (SELECT id FROM UF WHERE nome = 'Rio Grande do Norte')),
('Caicó', (SELECT id FROM UF WHERE nome = 'Rio Grande do Norte')),

-- Rio Grande do Sul
('Porto Alegre', (SELECT id FROM UF WHERE nome = 'Rio Grande do Sul')),
('Canoas', (SELECT id FROM UF WHERE nome = 'Rio Grande do Sul')),
('Pelotas', (SELECT id FROM UF WHERE nome = 'Rio Grande do Sul')),
('Santa Maria', (SELECT id FROM UF WHERE nome = 'Rio Grande do Sul')),

-- Rondônia
('Porto Velho', (SELECT id FROM UF WHERE nome = 'Rondônia')),
('Ji-Paraná', (SELECT id FROM UF WHERE nome = 'Rondônia')),
('Cacoal', (SELECT id FROM UF WHERE nome = 'Rondônia')),
('Rolim de Moura', (SELECT id FROM UF WHERE nome = 'Rondônia')),

-- Roraima
('Boa Vista', (SELECT id FROM UF WHERE nome = 'Roraima')),
('Rorainópolis', (SELECT id FROM UF WHERE nome = 'Roraima')),
('Caracaraí', (SELECT id FROM UF WHERE nome = 'Roraima')),
('Cantá', (SELECT id FROM UF WHERE nome = 'Roraima')),

-- Santa Catarina
('Florianópolis', (SELECT id FROM UF WHERE nome = 'Santa Catarina')),
('Joinville', (SELECT id FROM UF WHERE nome = 'Santa Catarina')),
('Blumenau', (SELECT id FROM UF WHERE nome = 'Santa Catarina')),
('Chapecó', (SELECT id FROM UF WHERE nome = 'Santa Catarina')),

-- São Paulo
('São Paulo', (SELECT id FROM UF WHERE nome = 'São Paulo')),
('Campinas', (SELECT id FROM UF WHERE nome = 'São Paulo')),
('São Bernardo do Campo', (SELECT id FROM UF WHERE nome = 'São Paulo')),
('Santo André', (SELECT id FROM UF WHERE nome = 'São Paulo')),

-- Sergipe
('Aracaju', (SELECT id FROM UF WHERE nome = 'Sergipe')),
('Lagarto', (SELECT id FROM UF WHERE nome = 'Sergipe')),
('Itabaiana', (SELECT id FROM UF WHERE nome = 'Sergipe')),
('Estância', (SELECT id FROM UF WHERE nome = 'Sergipe')),

-- Tocantins
('Palmas', (SELECT id FROM UF WHERE nome = 'Tocantins')),
('Araguaína', (SELECT id FROM UF WHERE nome = 'Tocantins')),
('Gurupi', (SELECT id FROM UF WHERE nome = 'Tocantins')),
('Tocantinópolis', (SELECT id FROM UF WHERE nome = 'Tocantins'));

INSERT INTO Cliente (nome, cpf, endereco, uf_id, cidade_id, ativo) VALUES 
('Maria Silva', '123.456.789-00', 'Rua A, 123, Centro', 1, 1, 'S'),
('João Pereira', '234.567.890-11', 'Avenida B, 456, Bairro', 1, 2, 'S'),
('Ana Costa', '345.678.901-22', 'Rua C, 789, Vila', 2, 1, 'N'),
('Pedro Santos', '456.789.012-33', 'Rua D, 321, Loteamento', 2, 3, 'S'),
('Carla Almeida', '567.890.123-44', 'Avenida E, 654, Parque', 1, 1, 'S');

INSERT INTO Corretor (nome, cpf) VALUES 
('Ricardo Gomes', '111.222.333-44'),
('Fernanda Lima', '222.333.444-55'),
('Lucas Martins', '333.444.555-66'),
('Juliana Rocha', '444.555.666-77'),
('Thiago Oliveira', '555.666.777-88');

INSERT INTO ClienteCorretor (corretor_id, cliente_id) VALUES 
(1, 1), 
(1, 2), 
(2, 3),  
(3, 4),  
(4, 5);  
