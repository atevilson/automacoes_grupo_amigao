# Automações Grupo Amigão

## Sumário Geral

1. [Projeto 1 - Robo Cobrança de Fornecedores sem agendamentos em centros de distribuições](#projeto1)  
   - [Objetivo](#objetivo1)  
   - [Tecnologias Utilizadas](#tecnologias-utilizadas1)  
   - [Como Funciona](#comofunciona1)  
   - [Observações Importantes](#observacoesimportantes1)  
2. [Projeto 2 - Automação de relatório carteira de pedidos](#projeto2)  
   - [Objetivo](#objetivo2)  
   - [Como Funciona](#comofunciona2)  
   - [Observações Importantes](#observacoesimportantes2)  
3. [Projeto 3 - Envio Rápido do Relatório de Pedidos por E-mail](#projeto3)  
   - [Objetivo](#objetivo3)  
   - [Como Funciona](#comofunciona3)  
   - [Observações Importantes](#observacoesimportantes3)  

---

## **Projeto 1** - Robo Cobrança de Fornecedores sem agendamentos em centros de distribuições <a name="projeto1"></a>

### 1. Objetivo <a name="objetivo1"></a>
O objetivo desta automação é identificar todos os pedidos ativos e sem data de agendamento (DT_AGENDA) no sistema, verificando também a classificação desses pedidos (apenas “Original” e com LOCAL_ENT = “CD”). Uma vez encontrados, são enviados e-mails de cobrança aos fornecedores responsáveis por tais pedidos. O principal ganho foi a velocidade do processo de cobrança que antes era feito de forma manual via planilha Excel pelo assistente, onde era enviado e-mail um a um para cada fornecedor anexando os pedidos no corpo do e-mail. O processo, que antes levava quase 1 semana (dada a quantidade de fornecedores), agora é feito automaticamente em massa em cerca de 10 minutos, variando apenas pela velocidade da rede e do servidor de e-mail.

### 2. Tecnologias Utilizadas <a name="tecnologias-utilizadas1"></a>
- **Python 3.7+**  
- **Bibliotecas**:
  - `pandas` para manipulação e filtragem de dados.
  - `datetime` para lidar com datas.
  - `win32com.client` (pywin32) para integração com o Outlook e envio de e-mails.

### 3. Como Funciona <a name="comofunciona1"></a>
1. **Carregamento das Bases**  
   - `base_dashboard.xlsx`: Concentra a carteira de pedidos com data de emissão, data de entrega, fornecedor etc.  
   - `emails_forn.xlsx`: Contém os endereços de e-mail dos fornecedores.  
   - `emails_amigao.xlsx`: Contém os endereços de e-mail da equipe interna.

2. **Filtragem e Análise**  
   - Seleciona pedidos com entrega >= 3 dias após a data atual.  
   - Verifica fornecedores com coluna DT_AGENDA iniciada em “SEM”.  
   - Filtra pedidos tipo “Original” para centros de distribuição (exceto 745 e 61).  
   - Remove duplicados pelo número do pedido.

3. **Envio de E-mail**  
   - Agrupa pedidos por (usuário, departamento, fornecedor).  
   - Localiza endereço principal (TO) e cópia (CC).  
   - Envia e-mail via Outlook listando cada pedido sem agendamento, com datas de emissão e prazo.

4. **Contadores**  
   - Apresenta no console o total de e-mails enviados e não enviados.

### 4. Observações Importantes <a name="observacoesimportantes1"></a>
- Se o fornecedor não tiver e-mail registrado, o envio é ignorado.  
- Em caso de falha, o script exibe o erro no console.  
- O Outlook precisa estar instalado e configurado.

---

## **Projeto 2** - Automação de relatório carteira de pedidos <a name="projeto2"></a>

### 1. Objetivo <a name="objetivo2"></a>
O objetivo inicial aqui é processar diversos relatórios TXT (r1 ao r6, cancelados, pendentes etc.), consolidá-los em um único DataFrame e remover pedidos que não atendem aos critérios (cancelados, pendentes, datas de previsão muito antigas). Ao final, gera-se uma planilha `base_dashboard.xlsx` com dados filtrados e tratados. O principal ganho com essa automação foi o tempo para gerá-lo com todo tratamento: antes era feito via planilha Excel e, devido ao grande volume de dados, o processo levava em torno de 2 horas. Com essa automação, conclui-se em menos de 2 minutos.

### 2. Como Funciona <a name="comofunciona2"></a>
1. **Processamento de TXT**  
   - Lê e valida cada arquivo `.txt` (31 colunas).  
   - Filtra linhas inconsistentes e salva em `temp.txt`.  
   - Converte para DataFrame e descarta duplicados.

2. **Integração com Informações Adicionais**  
   - Lê relatórios de pedidos cancelados e pendentes.  
   - Faz merges para remover cancelados (`St = C`) e pendentes (`Autorização = P`).  
   - Realiza diversas limpezas de colunas e formatações (datas, floats, etc.).

3. **Tratamentos Finais**  
   - Cria colunas como `Chave`, `Tipo Pedido`, `Classificacao Pedido`.  
   - Exclui duplicados pelo conjunto de colunas relevantes.  
   - Gera a planilha final `base_dashboard.xlsx`.

### 3. Observações Importantes <a name="observacoesimportantes2"></a>
- Verifique se o encoding de todos os arquivos `.txt` está correto (`iso-8859-1`).  
- Arquivos incompletos ou com formato divergente terão linhas descartadas.  
- O script remove o arquivo temporário `temp.txt` ao final para limpeza.

---

## **Projeto 3** - Envio Rápido do Relatório de Pedidos por E-mail <a name="projeto3"></a>

### 1. Objetivo <a name="objetivo3"></a>
O principal objetivo desta automação é **enviar rapidamente**, com apenas um clique, o relatório de compras (`PEDIDOS_COMPRA.xlsx`) a toda a equipe responsável, sem a necessidade de abrir o Outlook manualmente. Dessa forma, a produtividade aumenta e evitam-se possíveis atrasos ou falhas humanas no envio.

### 2. Como Funciona <a name="comofunciona3"></a>
1. **Carregamento de Arquivos**  
   - Carrega o `PEDIDOS_COMPRA.xlsx` e uma imagem de assinatura (caso necessário) para inserir no corpo do e-mail.

2. **Criação de E-mail Automatizado**  
   - Utiliza `win32com.client` para inicializar o Outlook e criar a mensagem.  
   - Define assunto, destinatários (To e CC) e corpo da mensagem em formato HTML.  
   - Insere a assinatura em Base64 diretamente no corpo do e-mail.

3. **Envio**  
   - Faz a anexação do arquivo `PEDIDOS_COMPRA.xlsx`.  
   - O e-mail é enviado automaticamente e, no console, é exibida a mensagem de sucesso.

### 3. Observações Importantes <a name="observacoesimportantes3"></a>
- Necessita de Outlook instalado e configurado.  
- A imagem de assinatura deve existir no caminho configurado no script.  
- A velocidade do envio depende da estabilidade da rede e do servidor de e-mail.

---

## Desenvolvedor

<sub><b>Atevilson Freitas</b> 🧑‍💻</sub>  
