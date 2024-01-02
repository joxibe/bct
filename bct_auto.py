import pandas as pd
from datetime import date, datetime
import os
import openpyxl


def main():
    input_file = 'bct.xlsx'
    output_folder = 'C:\\Users\\adminoperaciones\\Downloads\\KOBA\\INFORMES\\2024\\BCT\\ENERO'
    # output_folder = 'X:\Escritorio\BCT automatizacion'

    df = leer_archivo(input_file)
    df = agregar_filtros(df)
    df = reorganizar_columnas(df)
    exportar_datos(df, output_folder)


def leer_archivo(filename):
    print('Leyendo archivo')
    input_cols = [3, 4, 5, 6, 7, 10, 12]
    df = pd.read_excel(filename, sheet_name='Worksheet',
                       header=0, usecols=input_cols)
    return df


def agregar_filtros(df):
    print('Filtrando y eliminando duplicados de la fecha:')
    fecha_actual = date.today().strftime("%d/%m/%y")
    print(fecha_actual)
    df = df[df['FECHA_TRX'] == fecha_actual]
    df = df.drop_duplicates()
    return df


def reorganizar_columnas(df):
    # colocar la columna 12(indice 11) fecha_trx de primera
    cols = df.columns.tolist()
    cols_new = ['FECHA_TRX', 'COD_TIENDA',  'DIVISION',
                'REGIONAL', 'DIRECCION_IP', 'POS', 'NOMBRE_TIENDA']
    df = df[cols_new]
    return df


def exportar_datos(df, output_folder):
    print('Exportando archivo procesado')
    fecha_actual = date.today().strftime("%d/%m/%y")
    hora_actual = datetime.now().strftime("%H-%M")
    # print(hora_actual)
    dia = fecha_actual[0:2]
    bct = f'bct_{hora_actual}.xlsx'
    # print('linea....', bct)
    bct_final = os.path.join(output_folder, dia, bct)
    # print(bct_final)

    if not os.path.exists(os.path.abspath(bct_final)):
        os.makedirs(os.path.dirname(bct_final), exist_ok=True)

     # Exportar el DataFrame a un archivo Excel
    df.to_excel(bct_final, index=False, engine='openpyxl')

    # Abrir el archivo Excel y ajustar el ancho de las columnas A hasta G
    workbook = openpyxl.load_workbook(bct_final)
    worksheet = workbook.active
    for col in worksheet.iter_cols(min_col=1, max_col=7):
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width

    workbook.save(bct_final)


if __name__ == '__main__':
    main()
    input('\tProceso finalizado, presiona enter para salir ... ')
