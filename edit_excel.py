"""Script voor Waterschap Noorderzijlvest voor het bewerken van lab-formulieren."""
import environment
environment.activate()
from openpyxl import load_workbook, writer
import click
from pathlib import Path
from copy import copy

# %%
@click.command()
@click.option("--nummer", type=int, help=("Regel nummer in Excel waarin iets "
                                          "moet worden ingevoegd "
                                          "voorbeeld: --nummer 4"))
@click.option('--regel', type=str, help=("Regel die moet worden ingevoegd, "
                                         'gescheiden met "," en tussen ""'
                                         'voorbeeld: --regel "meetdoel,KRW"'))
def main(nummer, regel):
    """Invoeren van regels in Excel."""
    regel = regel.split(",")
    input_dir = Path("input").absolute()
    output_dir = Path("output").absolute()
    output_dir.mkdir(exist_ok=True)
    insert_number = nummer - 1
#    row_df = pd.DataFrame([regel])
    for excel_file in input_dir.glob("*.xlsx"):
        file_name = excel_file.name
        print(file_name)
        out_file = output_dir.joinpath(file_name)
        if out_file.exists():
            out_file.unlink()
        book = load_workbook(excel_file)
        ws = book.worksheets[0]
        ws.insert_rows(nummer)
        #invullen van de cellen, inclusief stijl
        for idx, cell in enumerate(ws[nummer]):
            cell.fill = copy(ws[nummer+1][idx].fill)
            cell.alignment = copy(ws[nummer+1][idx].alignment)
            cell.font = copy(ws[nummer+1][idx].font)
            cell.border = copy(ws[nummer+1][idx].border)
            if idx < len(regel):
                value = regel[idx]
            else:
                value = ""
            cell.value = value
        xls_writer = writer.excel.save_workbook(book, out_file)
#        xls_writer = pd.ExcelWriter(out_file, engine="openpyxl")
        # xls_writer.book = book
        # xls_writer.save()
        # df = pd.read_excel(excel_file, engine="openpyxl", header=None)
        # df = pd.concat([df.loc[0:insert_number-1],
        #                 row_df,
        #                 df.loc[insert_number:]],
        #                axis=0)
        # df.to_excel(out_file)


if __name__ == '__main__':
    main()
