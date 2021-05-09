import numpy as np
import pandas as pd
import xlsxwriter, os, re, sys

pictureCol = 0
tableCol = 5
spaceBetweenBooks = 4
tableCellWidth = 12

def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable) 
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

imagePath = resource_path("SampleCover.jpg")

config = {
    'ingram': {
        'tsv':True,
        'colsToKeep': ['author','title','MTD_Quantity','MTD_pub_comp','PTD_Quantity','PTD_pub_comp','reporting_currency_code'],
        'renamedCols': {
            'author': 'Author',
            'title': 'Title',
            'MTD_Quantity': 'Units',
            'MTD_pub_comp': 'Royalty',
            'PTD_Quantity': 'Units',
            'PTD_pub_comp': 'Royalty',
            'reporting_currency_code': 'Currency',
        }
    },
    'amazonKindle': {
        'tsv':False,
        'sheet': 'eBook Royalty',
        'colsToKeep': ['Title','Author Name','Net Units Sold','Royalty','Currency'],
        'renamedCols': {
            'Net Units Sold': 'Units',
            'Author Name': 'Author',
        }
    },
    'amazonPaper': {
        'tsv':False,
        'sheet': 'Paperback Royalty',
        'colsToKeep': ['Title','Author Name','Net Units Sold','Royalty','Currency'],
        'renamedCols': {
            'Net Units Sold': 'Units',
            'Author Name': 'Author',
        }
    },
}


def getBookData(df, author, title):
    if len(df) == 0:
        return pd.DataFrame(columns=['Units','Royalty','Currency'])
    df = df[df['Author'] == author]
    df = df[df['titleid'] == titleToTitleid(title)]
    df = df[['Units','Royalty','Currency']]
    # Get sums for each currency
    df = df.groupby(df['Currency'], as_index=False).sum()
    df = df.reset_index(drop=True)
    df = df[['Units','Royalty','Currency']] # Reorder
    return df

def getVendorTotals(df, rates):
    if len(df) == 0:
        return pd.DataFrame(columns=['Units','Royalty','Currency'])
    totalRoyalty = 0
    totalUnits = 0
    for _, row in df.iterrows():
        totalRoyalty = totalRoyalty + row['Royalty'] * rates[row['Currency']]
        totalUnits = totalUnits + row['Units']

    return pd.DataFrame({
        'Units':totalUnits,
        'Royalty':totalRoyalty,
        'Currency':'AUD'
    }, index=[0])

def getCleanDataFromPath(cfg, path):
    if path is None:
        # Return empty df
        return pd.DataFrame(columns=['Title','Author','Units','Royalty','Currency'])
    
    # Read file
    if cfg['tsv']:
        try:
            df = pd.read_csv(path,sep='\t',header=0,encoding='latin1')
        except:
            df = pd.read_excel(path)
    else:
        df = pd.read_excel(path,sheet_name=cfg['sheet'])

    # Remove unwanted columns
    for col in df.columns:
        if not col in cfg['colsToKeep']:
            del df[col]

    # Rename columns
    df.rename(cfg['renamedCols'],axis='columns',inplace=True)

    # Remove empty rows
    df.dropna(subset=["Title"], inplace=True)

    for index, row in df.iterrows():
        # Change Title to titleid
        titleid = titleToTitleid(row['Title'])
        df.loc[index,'titleid'] = titleid

        # Change Author from 'Last, First' to 'First Last'
        author = row['Author']
        author = author.split(', ')
        if len(author) > 1:
            author.reverse()
            author = ' '.join(author)
            df.loc[index,'Author'] = author

    return df

def titleToTitleid(title):
    return title.lower().replace(' ','').replace(':','')

def dfToExcel(writer, df, startrow, tableTitle, rates, sheet_name='Sheet1', totalRow=True):
    # Main table
    # df = df.round(2)
    df.to_excel(writer,sheet_name=sheet_name,index=False,startrow=startrow+1,startcol=tableCol)
    worksheet = writer.sheets[sheet_name]
    worksheet.write(startrow+1,tableCol,'Units')
    worksheet.write(startrow+1,tableCol+1,'Royalty')
    worksheet.write(startrow+1,tableCol+2,'Currency')

    workbook = writer.book
    title_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})

    # Title
    worksheet.merge_range(startrow,tableCol,startrow,tableCol+2,tableTitle.upper(),title_format)
    
    # Totals row
    currentRow = startrow + 3 + len(df)
    totals = getVendorTotals(df, rates).round(2)
    if (totalRow):
        worksheet.write(startrow+2+len(df),tableCol-1,'Total')
        totals.to_excel(writer,sheet_name=sheet_name,index=False,startrow=startrow+2+len(df),startcol=tableCol,header=False)
        currentRow = currentRow + 1

    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    worksheet.set_column(tableCol+1, tableCol+1, tableCellWidth, currency_format)

    return (currentRow, totals[['Units','Royalty']])

def bookToExcel(author, bookTitle, writer, startrow, ingramDf, amazonKindleDf, amazonPaperDf, rates, sheet_name='Sheet1'):
    authorsIngram = getBookData(ingramDf,author,bookTitle)
    authorsAmazonKindle = getBookData(amazonKindleDf,author,bookTitle)
    authorsAmazonPaper = getBookData(amazonPaperDf,author,bookTitle)

    currentRow = startrow + 2
    totals = pd.DataFrame([{'Units': 0, 'Royalty': 0}])
    for (df, tableTitle) in [(authorsIngram, 'INGRAMSPARK PAPERBACK'), (authorsAmazonKindle, 'AMAZON KDP KINDLE'), (authorsAmazonPaper, 'AMAZON KDP PAPERBACK')]:
        if len(df) > 0:
            (currentRow, thisTotal) = dfToExcel(writer=writer, sheet_name=sheet_name, df=df,  startrow=currentRow, tableTitle=tableTitle, totalRow=True, rates=rates)
            totals = totals + thisTotal

    bold_format = writer.book.add_format({'bold': True, 'align':'right'})
    worksheet = writer.sheets[sheet_name]
    worksheet.write(startrow,pictureCol,'Title:', bold_format)
    worksheet.write(startrow,pictureCol+1,bookTitle)
    worksheet.write(startrow+1,pictureCol,'Author:', bold_format)
    worksheet.write(startrow+1,pictureCol+1,author)
    worksheet.insert_image(startrow+2, 0, imagePath, {'x_offset': 26, 'y_offset': 10, 'x_scale': 0.65, 'y_scale': 0.65})

    totals['Currency'] = 'AUD'
    (currentRow, _) = dfToExcel(writer=writer, sheet_name=sheet_name, df=totals, startrow=currentRow, tableTitle="TOTAL SALES", totalRow=False, rates=rates)

    return max(currentRow,startrow+18)

def getAuthors(ingramDf, amazonKindleDf, amazonPaperDf):
    authors = set()
    for df in [ingramDf, amazonKindleDf, amazonPaperDf]:
        for author in df['Author']:
            authors.add(author)
    return authors

def getAuthorsTitles(author, ingramDf, amazonKindleDf, amazonPaperDf):
    titles = set()
    titleids = set()
    for df in [ingramDf, amazonKindleDf, amazonPaperDf]:
        for title in df[df['Author'] == author]['Title']:
            titleid = titleToTitleid(title)
            if titleid not in titleids:
                titleids.add(titleid)
                titles.add(title)

    return titles

def authorToExcel(author, ingramDf, amazonKindleDf, amazonPaperDf, periodTitle, outputPath, rates):
    writer = pd.ExcelWriter(f'{outputPath}/{author}.xlsx', engine='xlsxwriter')
    
    currentRow = 2
    for title in getAuthorsTitles(author,ingramDf, amazonKindleDf, amazonPaperDf):
        currentRow = bookToExcel(author, title, writer, currentRow, ingramDf, amazonKindleDf, amazonPaperDf, rates)
        currentRow = currentRow + spaceBetweenBooks

    title_format = writer.book.add_format({'bold': True, 'font_size':20, 'align':'right'})
    worksheet = writer.sheets['Sheet1']
    worksheet.write(0,tableCol+2,'CILENTO PUBLISHING SALES REPORT', title_format)
    worksheet.write(1,tableCol+2,periodTitle, title_format)
    worksheet.set_column(tableCol, tableCol+3, tableCellWidth)

    writer.save()

def main(ingramPaths, amazonPath, outputPath, rates):    
    if amazonPath is not None:
        periodTitle = re.search(r'(?<=Dashboard-).*(?=-\d{4}\.xlsx)', amazonPath).group(0)
    else:
        periodTitle = "'===> UNKNOWN MONTH <==="

    # Get Ingram
    ingramPaths = [p for p in ingramPaths if p is not None]
    ingramDfs = [getCleanDataFromPath(config['ingram'], p) for p in ingramPaths]
    ingramDf = pd.concat(ingramDfs, ignore_index=True)

    # Get Amazon
    amazonKindleDf = getCleanDataFromPath(config['amazonKindle'], amazonPath)
    amazonPaperDf = getCleanDataFromPath(config['amazonPaper'], amazonPath)

    for author in getAuthors(ingramDf, amazonKindleDf, amazonPaperDf):
        authorToExcel(author, ingramDf, amazonKindleDf, amazonPaperDf, periodTitle, outputPath, rates)

    return []

if __name__ == '__main__':
    # For testing purposes only
    ingramPaths = ['../data/2020-july/sales_compAU.xls','../data/2020-july/sales_compUK.xls']
    amazonPath = '../data/2020-july/KDP-Sales-Dashboard-JULY-2020.xlsx'
    outputPath = '../data/2020-july-output-new'
    rates = {'AUD':1,'USD':1.5,'GBP':2}

    main(ingramPaths, amazonPath, outputPath, rates)
