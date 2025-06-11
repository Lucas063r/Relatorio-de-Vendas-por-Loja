import pandas as pd
import win32com.client as win32

def carregar_dados(caminho_arquivo):
    try:
        return pd.read_excel(caminho_arquivo)
    except FileNotFoundError:
        print("Arquivo de vendas não encontrado.")
        exit()

def calcular_faturamento(df):
    return df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

def calcular_quantidade_vendida(df):
    return df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

def calcular_ticket_medio(faturamento, quantidade):
    return (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame(name='Ticket Médio')

def enviar_email(destinatario, faturamento, quantidade, ticket_medio):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.Subject = 'Relatório de Vendas por Loja'
    corpo_email = f"""
    <p>Prezados,</p>

    <p>Segue o relatório de Vendas por cada Loja.</p>

    <p><strong>Faturamento:</strong></p>
    {faturamento.to_html(formatters={{'Valor Final': 'R${{:,.2f}}'.format}})}

    <p><strong>Quantidade Vendida:</strong></p>
    {quantidade.to_html()}

    <p><strong>Ticket Médio dos Produtos em cada loja:</strong></p>
    {ticket_medio.to_html(formatters={{'Ticket Médio': 'R${{:,.2f}}'.format}})}

    <p>Qualquer dúvida estou à disposição.</p>

    <p>Att.,<br>Lucas Alves</p>
    """
    mail.HTMLBody = corpo_email
    mail.Send()

def main():
    df = carregar_dados("data/Vendas.xlsx")
    faturamento = calcular_faturamento(df)
    quantidade = calcular_quantidade_vendida(df)
    ticket_medio = calcular_ticket_medio(faturamento, quantidade)

    enviar_email(
        destinatario='seu_email@exemplo.com',
        faturamento=faturamento,
        quantidade=quantidade,
        ticket_medio=ticket_medio
    )

if __name__ == "__main__":
    main()
