import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import BooleanCondition
from gspread_formatting import format_cell_range

participantes = [
    "João Augusto", "Victor Melo", "Alan Holanda", "Heitor Guerra",
    "Eduardo Morais", "Pedro Silva", "Ewerton Carlos", "Gustavo C.", "Gustavo H."
]


meses = {
    "JUNHO": [11, 12],
    "JULHO": [13, 14, 15, 16, 17],
    "AGOSTO": [18, 19, 20, 21, 22],
    "SETEMBRO": [23, 24, 25],
    "OUTUBRO": [26, 27, 28],
    "NOVEMBRO": [29, 30, 31, 32],
    "DEZEMBRO": [33, 34, 35, 36, 37, 38]
}


scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

spreadsheet_id = "18aAiX2VtwYpbrTcblCCD3yvo8E4cf70DDMayoEc-v24"
spreadsheet = client.open_by_key(spreadsheet_id)


from gspread_formatting import (
    get_conditional_format_rules, ConditionalFormatRule,
    BooleanRule, CellFormat, Color
)
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import gspread.utils


def preparar_planilha(sheet, rodadas_do_mes):
    total_col = len(rodadas_do_mes) + 2  
    letra_total = gspread.utils.rowcol_to_a1(1, total_col)[0]  

    
    for i, rodada in enumerate(rodadas_do_mes):
        col = i + 2
        if not sheet.cell(1, col).value:
            sheet.update_cell(1, col, f"Rodada {rodada}")
   
    sheet.update_cell(1, total_col, "TOTAL")

    
    for idx, nome in enumerate(participantes, start=2):
        if sheet.cell(idx, 1).value != nome:
            sheet.update_cell(idx, 1, nome)
        letra_inicio = gspread.utils.rowcol_to_a1(1, 2)[0]
        letra_fim = gspread.utils.rowcol_to_a1(1, total_col - 1)[0]
        formula = f"=SUM({letra_inicio}{idx}:{letra_fim}{idx})"
        cell_addr = gspread.utils.rowcol_to_a1(idx, total_col)
        sheet.update_acell(cell_addr, formula)

    
    faixa = f"${letra_total}$2:${letra_total}${len(participantes)+1}"
    regra = ConditionalFormatRule(
        ranges=[{
            "sheetId": sheet._properties['sheetId'],
            "startRowIndex": 1,
            "endRowIndex": len(participantes) + 1,
            "startColumnIndex": total_col - 1,
            "endColumnIndex": total_col
        }],
        booleanRule=BooleanRule(
            condition=BooleanCondition('CUSTOM_FORMULA', [f'=ROW()=MATCH(MAX({faixa}), {faixa}, 0)+1']),
            format=CellFormat(backgroundColor=Color(0.8, 1, 0.8))  # verde claro
        )

    )

    
    rules = get_conditional_format_rules(sheet)
    rules.clear()  
    rules.append(regra)  
    rules.save()
    
    centralizado = CellFormat(horizontalAlignment='CENTER')

    
    letra_inicio = gspread.utils.rowcol_to_a1(1, 2)[0]  
    letra_total = gspread.utils.rowcol_to_a1(1, total_col)[0]  
    faixa_toda = f"{letra_inicio}2:{letra_total}{len(participantes)+1}"

    
    format_cell_range(sheet, faixa_toda, centralizado)


'''def aplicar_formatacao_total(sheet, total_col, participantes):
    from gspread_formatting import (
        get_conditional_format_rules, ConditionalFormatRule,
        BooleanRule, BooleanCondition, CellFormat, Color
    )

    letra_total = gspread.utils.rowcol_to_a1(1, total_col)[0]
    faixa = f"{letra_total}2:{letra_total}{len(participantes)+1}"

    regra = ConditionalFormatRule(
        ranges=[{
            "sheetId": sheet._properties['sheetId'],
            "startRowIndex": 1,
            "endRowIndex": len(participantes) + 1,
            "startColumnIndex": total_col - 1,
            "endColumnIndex": total_col
        }],
        booleanRule=BooleanRule(
            condition=BooleanCondition('CUSTOM_FORMULA', [
                f"={letra_total}1=MAX({faixa})"
            ]),
            format=CellFormat(backgroundColor=Color(0.8, 1, 0.8))  # Verde claro
        )
    )

    rules = get_conditional_format_rules(sheet)
    rules.clear()
    rules.append(regra)
    rules.save()'''


def mostrar_pontuacoes(sheet, rodada, rodadas_do_mes):
    print(f"\nRodada {rodada}")
    col = rodadas_do_mes.index(rodada) + 2  
    dados = sheet.get_all_values()
    for i, nome in enumerate(participantes, start=2):
        try:
            valor = dados[i - 1][col - 1]
        except IndexError:
            valor = ''
        print(f"{nome}: {valor if valor else 0}")

def inserir_pontuacoes(sheet, rodada, rodadas_do_mes, total_col):
    col = rodadas_do_mes.index(rodada) + 2
    for i, nome in enumerate(participantes, start=2):
        entrada = input(f"Pontuação de {nome}: ")
        if entrada.strip() != "":
            try:
                sheet.update_cell(i, col, float(entrada))
            except ValueError:
                print("Valor inválido, ignorado.")
                
    #aplicar_formatacao_total(sheet, total_col, participantes)
    pass

def alterar_pontuacao_individual(sheet, rodada, rodadas_do_mes, total_col):
    col = rodadas_do_mes.index(rodada) + 2
    print("\nParticipantes:")
    for idx, nome in enumerate(participantes, start=1):
        print(f"{idx} - {nome}")

    escolha = input("Digite o número do participante que deseja alterar: ")

    try:
        escolha = int(escolha)
        if 1 <= escolha <= len(participantes):
            i = escolha + 1
            nome = participantes[escolha - 1]
            atual = sheet.cell(i, col).value
            novo_valor = input(f"{nome} (atual: {atual if atual else 0}) -> Novo valor: ")
            if novo_valor.strip() != "":
                sheet.update_cell(i, col, float(novo_valor))
        else:
            print("Número inválido.")
    except ValueError:
        print("Entrada inválida.")
        
    #aplicar_formatacao_total(sheet, total_col, participantes)
    pass

def main():
    mes = input("Informe o mês (JUNHO, JULHO, AGOSTO, ...): ").strip().upper()
    if mes not in meses:
        print("Mês inválido.")
        return

    rodadas_do_mes = meses[mes]
    print(f"Rodadas disponíveis para {mes}: {rodadas_do_mes}")
    rodada = int(input("Escolha a rodada: "))
    if rodada not in rodadas_do_mes:
        print("Rodada não pertence ao mês escolhido.")
        return

    sheet = spreadsheet.worksheet(mes)
    preparar_planilha(sheet, rodadas_do_mes)
    mostrar_pontuacoes(sheet, rodada, rodadas_do_mes)

    while True:
        print("\n1 - Inserir pontuação para todos")
        print("2 - Alterar pontuação de um participante")
        print("3 - Sair")
        opcao = input("Escolha: ")
        
        total_col = len(rodadas_do_mes) + 2

        if opcao == '3':
            break
        elif opcao == '1':
            inserir_pontuacoes(sheet, rodada, rodadas_do_mes, total_col)
            mostrar_pontuacoes(sheet, rodada, rodadas_do_mes)
        elif opcao == '2':
            alterar_pontuacao_individual(sheet, rodada, rodadas_do_mes, total_col)
            mostrar_pontuacoes(sheet, rodada, rodadas_do_mes)
        else:
            print("Opção inválida.")

if __name__ == "__main__":
    main()
