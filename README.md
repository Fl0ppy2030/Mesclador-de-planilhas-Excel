# Junção Automática de Planilhas Excel

## 📌 Descrição

Este projeto foi criado para **automatizar a junção de duas planilhas do Excel**, eliminando um trabalho manual repetitivo. O script permite que o usuário selecione os arquivos por meio de uma interface gráfica simples e gera uma nova planilha consolidada como resultado.

A ferramenta é ideal para cenários em que duas planilhas possuem dados complementares (por exemplo, cadastro + informações adicionais) e precisam ser combinadas com frequência.

---

## 🛠️ Tecnologias Utilizadas

* **Python**
* **Pandas** – manipulação e tratamento de dados
* **Tkinter** – interface gráfica para seleção de arquivos e mensagens
* **OpenPyXL / Excel** – leitura e escrita de arquivos `.xlsx`

---

## ⚙️ Funcionamento Geral

1. O programa abre uma **janela gráfica** para o usuário.
2. O usuário seleciona:

   * A **primeira planilha Excel**
   * A **segunda planilha Excel**
3. O código:

   * Lê os arquivos utilizando o Pandas
   * Trata possíveis valores vazios ou incompatíveis
   * Realiza a **junção das planilhas** com base em uma coluna em comum
4. Um novo arquivo Excel é gerado com os dados consolidados.
5. Uma mensagem informa se o processo foi concluído com sucesso ou se ocorreu algum erro.

---

## 🔗 Lógica da Junção

* As planilhas são carregadas com `pandas.read_excel()`
* A junção é feita utilizando `pandas.merge()`


## 🧠 Tratamento de Erros

O código possui validações para:

* Arquivo não selecionado
* Arquivo inválido ou corrompido
* Colunas incompatíveis
* Conversão incorreta de tipos de dados

Em caso de erro, o usuário é notificado através de uma **messagebox**.

---

## 📂 Estrutura do Processo

```text
Usuário
  ↓
Seleciona planilhas (Tkinter)
  ↓
Leitura dos dados (Pandas)
  ↓
Tratamento e junção
  ↓
Geração do novo Excel
```

---

## 🚀 Benefícios

* Economia de tempo
* Redução de erros manuais
* Processo padronizado
* Fácil de usar, mesmo sem conhecimento técnico

---

## ▶️ Como Usar

1. Execute o script Python
2. Selecione as duas planilhas solicitadas
3. Aguarde o processamento
4. Abra o arquivo Excel gerado

---

## 📌 Observações

* As planilhas devem estar fechadas antes da execução
* Recomenda-se manter os nomes das colunas padronizados
* O script pode ser convertido em `.exe` usando **PyInstaller**

---


Projeto desenvolvido para automatizar tarefas repetitivas e facilitar o fluxo de trabalho com planilhas Excel.

---

Se necessário, o código pode ser facilmente adaptado para juntar mais planilhas ou aplicar filtros adicionais.





* README escrito pelo CHATGPT
