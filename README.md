# Projeto Controle de Máscaras N95/PFF

Este projeto foi criado para ajudar a controlar a distribuição de máscaras N95/PFF em um hospital durante a pandemia de COVID-19. O objetivo era evitar o desperdício de máscaras, que estavam em falta na época.

## Sobre o Projeto

O projeto começou como uma simples planilha do Excel, mas evoluiu para incluir funcionalidades de VBA. A interface inclui um relógio automático com a data atual e um campo de texto para inserir o CPF do funcionário que vai retirar a máscara.

Ao inserir o CPF, o sistema busca na planilha de cadastro e retorna os dados do CPF, que são: NOME, EMPRESA e FUNÇÃO. Ao lado desses dados, aparecem os dados da última retirada do CPF, que são: data da última retirada e o status (dentro ou fora do prazo).

Há também um campo para justificativa, com opções pré-definidas, que é usado se o profissional estiver pegando a máscara antes do prazo de 15 dias. Quando o cadastro é novo ou para um paciente, informamos "colaborador novo" e/ou "visitante" para indicar no relatório que a saída foi para terceiros e não para os profissionais do hospital.

A retirada é registrada em uma planilha que serve como banco de dados para as retiradas (Planilha 3 - Controle) e o projeto também tem uma função de salvar a cada registro para evitar erros e perdas de dados.

## Interface do Usuário

A interface do usuário foi projetada para ser intuitiva e fácil de usar. Aqui estão os principais componentes:

### Cabeçalho

O cabeçalho é verde e exibe os logotipos e o texto relacionados ao "Ministério da Saúde", "FIOCRUZ Fundação Oswaldo Cruz" e "INI Evandro Chagas". O título em letras grandes indica “CENTRO HOSPITALAR COVID-19” seguido pela indicação da localização e a data/hora.

### Ícones de Função

Abaixo das informações do cabeçalho, existem ícones para diferentes funções:

- **Relatórios**: Acesso aos relatórios gerados pelo sistema.
- **Gerenciamento**: Acesso às planilhas e aos arquivos VBA, protegido por senha.
- **Informativo**: Indica as mudanças no formulário de acordo com a versão. Atual 3.7.
- **Fechar**: Fecha o "sistema", acionando uma macro que salva o projeto para evitar perda de dados.

Antes do campo do CPF, existem mais ícones de função:

- **Limpar**: Apaga os dados da busca, caso desista de fazer a retirada ou apenas pesquisar o status do CPF para a retirada.
- **Cadastro Novo**: Só é acionado se o CPF for informado. O campo do CPF segue a lógica da Receita Federal, evitando o cadastro de CPF inválido e/ou errado.
- **Três Pontos**: Abre um outro formulário que permite a busca mais avançada, por nome, empresa e função.

Ao lado da justificativa, existe o botão de dispensa, para registrar a saída da máscara e gravar no banco de dados.

### Informações do Usuário

No canto inferior esquerdo, há um texto indicando que o usuário. Isso permite ao usuário confirmar que está conectado com a conta correta.

## Contribuições

Este é um projeto pessoal e não estou buscando contribuições no momento. No entanto, se você tiver alguma sugestão ou feedback, sinta-se à vontade para me enviar uma mensagem.

## Contato

Se você tiver alguma dúvida sobre este projeto, sinta-se à vontade para me enviar uma mensagem.

