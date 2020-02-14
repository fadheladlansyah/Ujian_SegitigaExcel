def segitigaExcel(kata):
    import xlsxwriter

    book = xlsxwriter.Workbook('segitigaExcel.xlsx')
    sheet = book.add_worksheet('Sheet1')
    
    kata = kata.replace(' ','')
    l = len(kata)
    # check does sentence allowed
    for i in range(1,l):
        l -= i
#         print(i,l)
        if l < 0:
            break
#         print(l)
        if l <= i:
            if l == 0:
#                 print('OK')
                # write excel for allowed sentence
                for row in range(i):
                    kata_p = kata[:row+1]
#                     print(kata_p)
                    for col, c in enumerate(kata_p):
#                         print(row, col, c)
                        sheet.write(row,col,c)
                    kata = kata[row+1:]
#                     print(kata)
                book.close()
            else :
                print('Mohon maaf, jumlah karakter tidak memenuhi syarat membentuk pola.')

# segitigaExcel('Purwadhika')
# segitigaExcel('Purwadhika Startup and Coding School @BSD')
# segitigaExcel('kode')
# segitigaExcel('kode python')
# segitigaExcel('Lintang')
