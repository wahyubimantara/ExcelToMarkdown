import pandas as pd
import numpy as np
import re

df = pd.read_excel(r'Book1.xlsx')
data = np.array(df)
rows, cols = np.shape(data)
# print('\n', data, '\n')

# Global var
Tanggal = "XXX"
NomorBukti = "XXX"
KodeRekening = "XXX"
Debit = "XXX"
Kredit = "XXX"

a = np.array(['xxx', ''])
b = np.array(['', 'xxx'])

PART1 = pd.DataFrame(
    {
        'Tanggal':[Tanggal],
        'Nomor Bukti':[NomorBukti]    
    }
)
# -----------------------------------

for row in range(rows):
    
    result = []
    pick = data[row]
    
    # split dan simpan ke variable
    Nama = pick[0]
    Sistem_Akuntansi = pick[1]
    Kasus = pick[2]
    Uraian_Kasus = pick[3]
    Jenis_Jurnal = pick[4]
    D = pick[5:8]
    K = pick[9:]
    # ----------------------------

    # Rebuild bentuk tabel
    for deb in D:
        if deb is not np.nan:
            result.append(np.concatenate((KodeRekening, deb, a), axis=None))
            
    for kred in K:
        if kred is not np.nan:
            result.append(np.concatenate((KodeRekening, kred, b), axis=None))
            
    # Convert array result ke dataframe  
    PART2 = pd.DataFrame(result, columns=[ 'Kode Rekening', 'Uraian', 'Debit', 'Kredit'])
    tabel = pd.concat([PART1, PART2], axis=1)
    # print('\n', tabel)
    # ----------------------------
    
    # Print dan convert ke markdown
    md = tabel.to_markdown(index=False)
    # print('\n',md, '\n')
    # ----------------------------
    
    # Modify md sesuaikan dengan format Tabel Anda
    md = re.sub("---*", "--", md)
    md = re.sub(":--", "---", md)
    md = re.sub("--:", "---", md)
    md = re.sub("nan", "", md)
    md = re.sub("nan", "   ", md)
    md = re.sub("  *", " ", md)
    # print('\n',md, '\n')
    # ----------------------------
    
    # Save to markdown format
    filename = str(row+1) + "."+ Sistem_Akuntansi +"-"+Kasus + '.md'
    with open(filename, "a") as f:
        print("---", file=f)
        print("sidebar_position: ",row+1, file=f)
        print("---", file=f)
        print("## "+Kasus,'\n', file=f)    
        print(":::tip Pejelasan",'\n', file=f)
        print (Uraian_Kasus,'\n', file=f)
        print(":::",'\n', file=f)
        print ('**',Jenis_Jurnal,'**',sep='', file=f)
        print('\n',md, '\n', file=f)
    # ---------------------------
