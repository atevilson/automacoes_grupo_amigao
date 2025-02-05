# README - Automação Grupo Amigão

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

---

## **Projeto 1** - Robo Cobrança de Fornecedores sem agendamentos em centros de distribuições <a name="projeto1"></a>

### 1. Objetivo <a name="objetivo1"></a>
O objetivo desta automação é identificar todos os pedidos ativos e sem data de agendamento (DT_AGENDA) no sistema, verificando também a classificação desses pedidos (apenas “Original” e com LOCAL_ENT = “CD”). Uma vez encontrados, são enviados e-mails de cobrança aos fornecedores responsáveis por tais pedidos.

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
O objetivo inicial aqui é processar diversos relatórios TXT (r1 ao r6, cancelados, pendentes etc.), consolidá-los em um único DataFrame e remover pedidos que não atendem aos critérios (cancelados, pendentes, datas de previsão muito antigas). Ao final, gera-se uma planilha `base_dashboard.xlsx` com dados filtrados e tratados. O principal ganho com essa automação foi o tempo para gera-lo com todo tratamento, antes era feito via planilha excel e devido ao grande volume de dados o tempo para conclusão do relatório era em torno de 2h, com essa automação o processo é feito em menos de 2 minutos.

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

## Desenvolvedor

<sub><b>Atevilson Freitas</b> 🧑‍💻</sub>