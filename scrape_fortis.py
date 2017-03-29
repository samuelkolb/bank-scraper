import pandas as pd
import datetime as dt





# fortis
# In_HEADER = Volgnummer;Uitvoeringsdatum;Valutadatum;Bedrag;Valuta rekening;Details;Rekeningnummer
# OU_HEADER = Name; Category; Details; Amount; Date; Bank

import xlsxwriter as xlsxwriter


# TODO hoe zit het met spaarrekeningen?

def import_fortis(path):
    csv_file = open(path, 'rb')

    df = pd.read_csv(csv_file, ';', decimal=',', encoding='utf-8', thousands='.')
    df = df.drop(['Volgnummer', 'Valutadatum', 'Valuta rekening', 'Rekeningnummer'], axis=1)
    cols = df.columns.tolist()
    cols = [cols[2]] + [cols[1]] + [cols[0]]
    df = df[cols]
    df.columns = ['Details', 'Amount', 'Date']

    df.insert(0, 'Name', '')
    df.insert(1, 'Category', '')
    df['Bank'] = '[FORTIS]'
    df['Date'] = pd.to_datetime(df.Date, format='%d/%m/%Y')

    return df

def import_ing_scrape(path):
    csv_file = open(path, 'rb')
    df = pd.read_csv(csv_file, ';', encoding='utf-8')

    df.columns = ['reknr', 'Details', 'Amount', 'Date']
    df = df.drop(['reknr'], axis=1)
    df.Amount = df.Amount.str.replace(' EUR', '')

    df.insert(0, 'Name', '')
    df.insert(1, 'Category', '')
    df['Bank'] = '[ING]'
    df['Date'] = pd.to_datetime(df.Date, format='%d/%m/%Y')

    return df


def import_ing(path):
    csv_file = open(path, 'rb')
    df = pd.read_csv(csv_file, ';', decimal=',', encoding='utf-8', thousands='.')

    df.insert(0, 'Name', '')
    df.insert(1, 'Category', '')

    df.insert(2, 'Details', '')
    df['Omschrijving'] = df['Omschrijving'].str.replace('\s\s+', ' ').str.replace(u'\ufffd', '_')
    df['Detail van de omzet'] = df['Detail van de omzet'].str.replace('\s\s+', ' ').str.replace(u'\ufffd', '_')
    df['Details'] = df['Omschrijving'].map(str) + ' [details vd omzet:] ' + df['Detail van de omzet'].map(str)

    df.insert(3, 'Amount', '')
    df['Amount'] = df['Bedrag']
    df.insert(4, 'Date', '')
    df['Date'] = pd.to_datetime(df.Valutadatum, format='%d/%m/%Y')
    df.insert(5, 'Bank', '')
    df['Bank'] = '[ING]'

    df = df.iloc[:, 0:6]
    return df


def import_revolut(path):
    csv_file = open(path, 'rb')
    df = pd.read_csv(csv_file, ',', encoding='utf-8')

    df.columns = ['Date', 'Details', 'Amount', 'a', 'b', 'c', 'd']
    df = df.drop(['a', 'b', 'c', 'd'], axis=1)
    cols = df.columns.tolist()
    cols = cols[1:] + [cols[0]]
    df = df[cols]


    df.insert(0, 'Name', '')
    df.insert(1, 'Category', '')
    df['Bank'] = '[REVOLUT]'
    df['Date'] = pd.to_datetime(df.Date, format='%d %b %Y ')

    return df


def write_months_to_excel(months, path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    workbook = writer.book

    for m in months:
        month = dt.datetime.strftime(m['Date'].iloc[0], '%b%y')
        m.Date = m.Date.dt.strftime('%d/%m%/%Y')
        m.to_excel(writer, month, float_format='%.2f', index=False)

        worksheet = writer.sheets[month]
        for row_num in xrange(2, 100):
            index_b = 'B' + str(row_num)
            index_g = 'G' + str(row_num)
            worksheet.data_validation(index_b, {'validate': 'list',
                                         'source': 'INDIRECT(' + index_g + ')'})

    write_categories(path, writer)
    writer.save()


def write_categories(path, writer):
    workbook = writer.book
    worksheet = workbook.add_worksheet('Categories')

    inkomsten = ('Gift', 'Investering', 'Loon', 'Overige', 'Zakgeld')
    uitgaven = ('Cash Afgehaald', 'Eten - Thuis', 'Eten - Uit', 'Gift', 'Herstelling', 'Kapper', 'Kleding',
                'Medisch', 'Onderdak - Huur', 'Onderdak - Uit', 'Terugkrijgen anderen', 'Terugkrijgen ouders',
                'Transport - Naft', 'Transport - OV', 'Transport - Werk', 'Uit - Drank', 'Uit - Inkom',
                'Uit - Overige', 'Utilities - Elektriciteit', 'Utilities - GSM', 'Utilities - Internet', 'Utilities - Water')


    # Write the data to a sequence of cells.
    worksheet.write('A1', 'Inkomst')
    worksheet.write('B1', 'Uitgave')
    worksheet.write_column('A2', inkomsten)
    worksheet.write_column('B2', uitgaven)
    workbook.define_name('Inkomst', '=Categories!A2:$A6')
    workbook.define_name('Uitgave', '=Categories!B2:$B23')



# ing = import_ing_scrape('csv/ing_scrape.csv')
ing = import_ing('csv/ing.csv')
fortis = import_fortis('csv/fortis.csv')
revolut = import_revolut('csv/revolut.csv')
out = pd.concat([ing, fortis, revolut])
out = out.sort_values(by='Date')

out['Amount'] = pd.to_numeric(out['Amount'])
out.insert(6, 'Type', '')
out['Type'] = out['Amount'].apply(lambda x: 'Uitgave' if x <= 0 else 'Inkomst')

months = []
for group in out.groupby(pd.Grouper(key='Date', freq='M')):
    months.append(group[1])

write_months_to_excel(months, 'test.xlsx')




# print out

