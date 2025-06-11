# 📊 Relatório de Vendas por Loja

Este projeto realiza uma análise automatizada de vendas utilizando Python e envia um relatório formatado por e-mail usando o Outlook.

---

## ✅ Funcionalidades

- Leitura de planilha Excel contendo dados de vendas
- Cálculo de **faturamento por loja**
- Cálculo da **quantidade total de produtos vendidos** por loja
- Cálculo do **ticket médio** por loja
- Envio automático do relatório por e-mail com **tabelas HTML formatadas**

---

## 🧱 Estrutura do Projeto

```
relatorio-vendas-por-loja/
├── data/
│   └── Vendas.xlsx            # Planilha de dados de entrada
├── src/
│   └── app.py                 # Script principal com a lógica da análise e envio do e-mail
├── requirements.txt           # Lista de bibliotecas necessárias
└── README.md                  # Este arquivo
```

---

## 🚀 Como Executar

1. **Clone este repositório**
```bash
git clone https://github.com/seu-usuario/relatorio-vendas-por-loja.git
cd relatorio-vendas-por-loja
```

2. **Instale as dependências**
```bash
pip install -r requirements.txt
```

3. **Adicione sua planilha**
> Substitua o arquivo `Vendas.xlsx` dentro da pasta `data/` pela sua própria planilha com colunas como `ID Loja`, `Valor Final` e `Quantidade`.

4. **Configure seu e-mail**
Abra o arquivo `src/app.py` e edite a linha:
```python
destinatario='seu_email@exemplo.com'
```

5. **Execute o projeto**
```bash
python src/app.py
```

> ⚠️ É necessário ter o **Outlook instalado e logado** na máquina para o envio de e-mail funcionar corretamente.

---

## 🧠 Tecnologias Utilizadas

- [Python 3](https://www.python.org/)
- [Pandas](https://pandas.pydata.org/)
- [PyWin32](https://pypi.org/project/pywin32/) para automação do Outlook
- [OpenPyXL](https://pypi.org/project/openpyxl/) para leitura de planilhas

---

## 🛠 Possíveis Erros

| Erro                         | Causa/Solução                                                               |
|------------------------------|------------------------------------------------------------------------------|
| `FileNotFoundError`          | O arquivo `Vendas.xlsx` não foi encontrado. Verifique o caminho.            |
| `ImportError`                | Falta alguma biblioteca. Use `pip install -r requirements.txt`              |
| Outlook não envia o e-mail   | Certifique-se de que o Outlook esteja instalado, aberto e configurado.      |
| `AttributeError`             | Verifique o uso de `.to_frame()` ao criar novas tabelas com uma única coluna |

---

## 📌 Próximos Passos (To-do)

- ✅ Refatorar código para uso de funções reutilizáveis
- ⬜ Adicionar tratamento de exceções mais robusto
- ⬜ Implementar testes com `pytest`
- ⬜ Criar interface gráfica simples com `tkinter` (futuro)

---

## 👨‍💻 Autor

Desenvolvido por **Lucas Alves**  
📧 la465551@gmail.com

---
