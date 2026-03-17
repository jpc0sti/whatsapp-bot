# Automação de WhatsApp: Criação de Grupos e Envio de Arquivos

## O que o projeto faz
Este projeto é uma automação construída em **Node.js** utilizando a biblioteca `whatsapp-web.js`. Ele atua como uma ponte entre uma planilha de controle no Excel e o WhatsApp. 

O script executa o seguinte fluxo de trabalho:
1. Lê a planilha `grupos.xlsm` (especificamente a aba `GrupoW`).
2. Identifica os códigos de clientes/grupos listados.
3. Acessa o diretório de rede mapeado (`Z:\Diversos\Processos\Topografia\Orçamentos\2025`) buscando a pasta correspondente ao código.
4. Cria um grupo no WhatsApp adicionando participantes predefinidos.
5. Filtra os arquivos válidos na pasta do cliente (ignorando subpastas e arquivos temporários).
6. Envia os documentos e imagens aprovados diretamente para o grupo recém-criado.
7. Gera um arquivo `status.txt` para sinalizar a conclusão ou erro para integrações externas (como macros VBA no Excel).

## Por que o projeto é útil
No setor de Topografia e Orçamentos, a criação de grupos de comunicação e o envio de documentação padrão para cada novo projeto é uma tarefa repetitiva e suscetível a falhas humanas (como esquecer de anexar um arquivo ou errar o número dos participantes). 

Esta automação elimina 100% do esforço manual dessa etapa. O que antes levaria minutos de cliques repetitivos por funcionários, agora é feito em segundos, rodando em segundo plano, garantindo padronização, agilidade na comunicação e rastreabilidade (com alertas de sucesso enviados ao administrador).

## Como os usuários podem começar a usar o projeto

### 1. Pré-requisitos
Certifique-se de que a máquina onde o robô vai rodar possui:
* **Node.js** instalado (versão 14+ recomendada).
* Mapeamento de rede ativo na unidade `Z:\` com as permissões de leitura adequadas.
* O arquivo `grupos.xlsm` na mesma pasta do script, com a aba `GrupoW` devidamente preenchida.
* Um celular com WhatsApp conectado à internet para leitura inicial do QR Code.

### 2. Instalação
Clone ou copie a pasta do projeto para o seu computador. Abra o terminal na pasta do projeto e instale as dependências executando:
```bash
npm install whatsapp-web.js qrcode-terminal xlsx

### 3. Execução Rápida (Windows)
Para usuários que preferem não utilizar o terminal, basta dar um duplo clique no arquivo `executar_robo.bat`. 
* O script verificará automaticamente se as dependências estão instaladas.
* O script iniciará a conexão com o WhatsApp.

### 4. Execução via Excel
Como operar a base de dados:
1. Abra o arquivo Excel modelo que acompanha esta pasta.
2. No painel de controle, selecione na lista o cliente para o qual deseja criar o grupo e enviar os arquivos.
3. Clique no botão "Confirmar" para disparar o robô de automação.
4. Aguarde a mensagem do formulário informando que a operação foi concluída com sucesso.