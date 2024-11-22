# Projeto Planilhas
Este projeto visa a automação do uso de planilhas usando Python e inteligência artificial. 

# Descrição
Este projeto possui duas versões, uma sendo a que usa o modelo Groq para a análise de dados com IA e a segunda que cuida dos dados de forma manual com entrada do usuário.

O modelo Groq foi selecionado por que no momento ele é o modelo mais fácil e prático de ser usado. Outra característica é que ele é o único modelo com uma chave de API de graça, já que outros modelo tipo ChatGPT da OpenAI e Claude da Anthropic pedem um pagamento inicial para usar suas chaves.

# Funcionalidades

### GPT
- Expressão matemática
- Análise de dados
- Descrepâncias

### Manual
- Inserir dados
- Modificar dados
- Escanear dados
- Criar planilha

As funcionalidades com GPT são apenas prompts customizados a partir do que o usuário selecionou e adicionou. Exemplo: se o usuário inseriu a opção de "Análise de dados" e especificou o que deve ser analizado, o prompt terá a parte de analização de dados e o que deve ser analizado, juntamente da planilha formatada para ser entendida pelo GPT.

As funcionalidades manuais são mais simples, elas permitem edição limitada e simples da planilha. Onde o usuário pode mudar um valor específico, tipo modificar onde existe o valor X por valor Y. Mas a parte principal deste programa é automação, onde o usuário pode inserir varios dados para cada coluna e linha de forma automatizada.

# Instalar e usar
Para instalar e usar apenas baixe o projeto direto da página do Github e extraia a pasta para onde quiser. 

### Dependências
Python (Qualquer versão a partir da 3)<br>
Groq 0.12.0 <br>
openpyxl 3.1.5 <br>
pandas 2.2.3 <br>
XlsxWriter 3.2.0 <br>

Para instalar as depedências é necessário que você tenha Python instalado no seu computador pois ele é o motor principal para o projeto e também por que pip vem junto com sua instalação.


##### 1. Baixe o projeto e o extraia onde quiser.

##### 2. Baixe e instale Python:
 https://www.python.org/downloads/
 
##### 3. Instale as dependências com pip:
Abra uma janela de terminal e mude o diretório atual para onde está o projeto extraído:

##### Exemplo no Windows e Linux:
```cd C:\Users\NOME_USUARIO\Documents\ProjetoPlanilhas ```<br>

##### Instale as dependências:
```pip install -r requirements.txt ```

##### 4. Execute o script:
```pyhton main.py``` <br>

###### Caso isso não funcione, apenas dê duplo click no main.py

# Como usar com GPT

### Pegue sua API Key do Groq:
##### 1. Acesse o site: https://console.groq.com/playground <br>
O site irá pedir seu login, então entre na sua conta ou crie caso não tenha uma.

##### 2. Na aba da esquerda, clique em API Keys
Clique no botão "Create API Key", então escolha um nome. Quando criado, copie sua chave e coloque um lugar para fácil acesso.

##### 3. Na execução do programa:
Após selecionar o uso de controlar as planilhas por GPT, cole e envie a sua API Key no console quando requisitado.
