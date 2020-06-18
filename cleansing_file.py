import editdistance
import pandas as pd

tempdict = {}
fname = '90.xlsx'
search_name = ['IMPORTER-QX', 'EXPORTER-QX']
record_book = {}
outname = fname + '_cldiff.xlsx'
out_book = {}


def initial(ind):
    record_book[ind] = {}
    tempdict[ind] = []
    out_book[ind] = pd.DataFrame(columns=["Distance", "A", "B", "Choice"])


def search_row(ind):
    if row[ind] not in tempdict[ind]:
        for j in tempdict[ind]:
            if (j, row[ind]) not in record_book[ind]:
                if (row[ind], j) in record_book[ind]:
                    record_book[ind][(j, row[ind])] = record_book[ind][(row[ind], j)]
                else:
                    score = editdistance.eval(j, row[ind])
                    record_book[ind][(j, row[ind])] = score
            else:
                score = record_book[ind][(j, row[ind])]
            if score < 7:
                out_book[ind] = out_book[ind].append(
                    pd.DataFrame({'Distance': [score], 'A': [j], 'B': [row[ind]], 'Choice': ['']}),
                    ignore_index=True)
        tempdict[ind].append(row[ind])


try:
    in_exc = pd.read_excel(fname, sheet_name=None)
except:
    print('cannot find the file')
    exit()
else:
    print('Read success')
in_df = in_exc['Trade Data']

for ind in search_name:
    initial(ind)

for index, row in in_df.iterrows():
    if index % 100 == 0:
        print(index)
    for ind in search_name:
        search_row(ind)

write = pd.ExcelWriter(outname)
for i in out_book:
    out_book[i].to_excel(excel_writer=write, sheet_name=i)
write.save()
write.close()
