# Robo Cobrança de Fornecedores sem agendamentos em centros de distribuições

Este projeto automatiza o envio de e-mails de cobrança de agendamentos de pedidos para cada fornecedor que possua pedidos sem data de agendamento registrada no relatório da carteira de pedidos.

---

## Sumário

1. [Objetivo](#objetivo)  
2. [Tecnologias Utilizadas](#tecnologias-utilizadas)  
3. [Como Funciona](#como-funciona)  
4. [Observações Importantes](#observacoes-importantes)  

---

### 1. Objetivo <a name="objetivo"></a>

O objetivo desta automação consistia em identificar todos os pedidos ativos e sem data de agendamento (DT_AGENDA) no sistema, verificando também a classificação desses pedidos (apenas “Original” e com LOCAL_ENT = “CD”). Uma vez encontrados, são enviados e-mails de cobrança aos fornecedores responsáveis por tais pedidos.

---

### 2. Tecnologias Utilizadas <a name="tecnologias-utilizadas"></a>

- **Python 3.7+**  
- **Bibliotecas**:
  - `pandas` para manipulação e filtragem de dados.
  - `datetime` para lidar com datas.
  - `win32com.client` (pywin32) para integração com o Outlook e envio de e-mails.

---

### 3. Como Funciona <a name="como-funciona"></a>

1. **Carregamento das Bases**:  
   São carregadas três planilhas:
   - `base_dashboard.xlsx` — base do relatório automatizado carteira de pedidos com datas de emissão, datas de entrega, classificação do pedido, fornecedor, etc.
   - `emails_forn.xlsx` — base com os endereços de e-mail dos fornecedores.
   - `emails_amigao.xlsx` — base com os endereços de e-mail da equipe interna (amigão) responsáveis pelo departamento.

2. **Filtragem e Análise**:
   - É filtrado o dataframe para considerar apenas pedidos cuja data de entrega seja maior ou igual a três dias após a data atual.
   - Identifica-se quais fornecedores estão com pedidos sem agendamento (DT_AGENDA começando com “SEM”).
   - Filtra-se ainda os pedidos do tipo “Original” e local de entrega “CD”, excluindo os CDs de número 745 e 61.
   - Ao final, remove duplicados pelo número de pedido.

3. **Envio de E-mail**:
   - Percorre-se cada grupo de pedidos por (usuário, departamento, fornecedor).
   - Encontra-se o e-mail do fornecedor no arquivo `emails_forn.xlsx`.
   - Cria-se a lista de cópia (`CC`) com base no mesmo arquivo de fornecedores e, se for encontrado o departamento do pedido no arquivo `emails_amigao.xlsx`, os contatos internos também são incluídos em cópia.
   - Para cada pedido sem agendamento, é construída a mensagem informando número do pedido, data de emissão e prazo para agendar.
   - O e-mail é enviado pelo Outlook, e o script registra via log simples se houve sucesso ou falha no envio.

4. **Contadores**:
   - São exibidos no console os totais de e-mails enviados com sucesso e e-mails que não puderam ser enviados.


### 4. Observações Importantes <a name="#observacoes-importantes"></a>
Em caso de falha no envio de e-mails, o script exibe o erro no console.
Caso um fornecedor não possua e-mail registrado em emails_forn.xlsx, o envio é ignorado para aquele fornecedor.
Verifique se o Outlook está instalado e ativo para o envio automático funcionar.

### Desenvolvedor
---

 <sub><b>Atevilson Freitas</b></sub></a> <a href="">🧑‍💻</a>