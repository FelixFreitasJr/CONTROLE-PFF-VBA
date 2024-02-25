# Projeto Controle de Máscaras N95/PFF
Este projeto foi criado para ajudar a controlar a distribuição de máscaras N95/PFF em um hospital durante a pandemia de COVID-19. O objetivo era evitar o desperdício de máscaras, que estavam em falta na época.

# Sobre o Projeto
O projeto começou como uma simples planilha do Excel, mas evoluiu para incluir funcionalidades de VBA. A interface inclui um relógio automático com a data atual e um campo de texto para inserir o CPF do funcionário que vai retirar a máscara.

Ao inserir o CPF, o sistema busca na planilha de cadastro e retorna os dados do CPF, que são: NOME, EMPRESA e FUNÇÃO. Ao lado desses dados, aparecem os dados da última retirada do CPF, que são: data da última retirada e o status (dentro ou fora do prazo).

Há também um campo para justificativa, com opções pré-definidas, que é usado se o profissional estiver pegando a máscara antes do prazo de 15 dias. Quando o cadastro é novo ou para um paciente, informamos “colaborador novo” e/ou “visitante” para indicar no relatório que a saída foi para terceiros e não para os profissionais do hospital.

A retirada é registrada em uma planilha que serve como banco de dados para as retiradas (Planilha 3 - Controle) e o projeto também tem uma função de salvar a cada registro para evitar erros e perdas de dados.

# Como Usar
Para usar este projeto, você precisará ter o Microsoft Excel instalado em seu computador. Abra o arquivo do projeto no Excel e siga as instruções na interface.

# Contribuições
Este é um projeto pessoal e não estou buscando contribuições no momento. No entanto, se você tiver alguma sugestão ou feedback, sinta-se à vontade para me enviar uma mensagem.

# Contato
Se você tiver alguma dúvida sobre este projeto, sinta-se à vontade para me enviar uma mensagem.
