import pandas as pd
import datetime

data = pd.read_excel("ruzgar_sayfa5.xlsx")

dd = pd.DataFrame(data)
#dd['YIL'] == 2009 or dd['YIL'] == 2010 or dd['YIL'] == 2011 or dd['YIL'] == 2012 or dd['YIL'] == 2013 or dd['YIL'] == 2014 or dd['YIL'] == 2015 or dd['YIL'] == 2016 or dd['YIL'] == 2017 or dd['YIL'] == 2018 or dd['YIL'] == 2019 or dd['YIL'] == 2020
#(dd[dd["Istasyon_No"]==sayi]).all()

while(True):


    istasyon = dd["Istasyon_No"]
    ruzgar = dd["GUNLUK_ORTALAMA_HIZI_m_sn"]
    dd['Tarih']=dd['YIL'].astype(str) + dd['AY'].astype(str).str.zfill(2)+ dd['GUN'].astype(str).str.zfill(2)
    tarih = dd['Tarihh'] = pd.to_datetime(dd['Tarih'], format='%Y%m%d')
    #tarih = dd[['Tarih']]=dd['YIL'].map(str) + '-' + dd['AY'].map(str) + '-' + dd['GUN'].map(str)
    aylar = dd["AY"]
    yillar = dd["YIL"]

    filename = 'islem1_syf5.xlsx'
    writer = pd.ExcelWriter( filename, engine='xlsxwriter')

    istasyon.to_excel(writer, sheet_name='Sheet1',startrow=0,startcol=0, index=False)
    tarih.to_excel(writer, sheet_name='Sheet1',startrow=0,startcol=1, index=False)
    ruzgar.to_excel(writer, sheet_name='Sheet1',startrow=0,startcol=2, index=False)
    aylar.to_excel(writer, sheet_name='Sheet1',startrow=0,startcol=3, index=False)
    yillar.to_excel(writer, sheet_name='Sheet1',startrow=0,startcol=4, index=False)


    writer.save()
