import pandas as pd
import openpyxl
import os,sys
from docxtpl import DocxTemplate
from docx2pdf import convert
import datetime
from openpyxl import Workbook
from xlwt import Workbook as Wbk



def imie_i_nazwisko(str):
    if len(str.split(' '))==3:
        return (str.split(' ')[0]).capitalize(),(str.split(' ')[2]).capitalize()
    return (str.split(' ')[0]).capitalize(),(str.split(' ')[1]).capitalize()

def kod_pocztowy(kod):
    if len(kod)==5:
        return str(kod[:2]+"-"+kod[2:])
    return kod

def utworzXlsDoImportu(lista_adresatow):
    do_importu_file = pd.ExcelFile('adresaci.xls')
    sheets = do_importu_file.sheet_names
    filename = datetime.datetime.now(datetime.timezone.utc).strftime("%d-%m-%Y %H %M")
    filepath = 'pisma_sądowe/'+str(dzisiejsza_data)+'/import/' + filename +'.xls'
    if os.path.exists('pisma_sądowe') == False:
        os.mkdir("pisma_sądowe", 0o666)
    if os.path.exists("pisma_sądowe/" + str(dzisiejsza_data)) == False:
        os.mkdir(("pisma_sądowe/" + str(dzisiejsza_data)), 0o666)
    if os.path.exists("pisma_sądowe/" + str(dzisiejsza_data)+'/import') == False:
        os.mkdir(("pisma_sądowe/" + str(dzisiejsza_data))+'/import', 0o666)
    book = Wbk(filepath)
    book.add_sheet(sheets[0])
    book.save(filepath)
    to_import = pd.read_excel('adresaci.xls', sheets[0])
    for i in lista_adresatow:
        to_import.loc[lista_adresatow.index(i), "AdresatNazwa"] = i[0]
        to_import.loc[lista_adresatow.index(i), "AdresatUlica"] = i[1]
        to_import.loc[lista_adresatow.index(i), "AdresatNumerDomu"] = i[2]
        to_import.loc[lista_adresatow.index(i), "AdresatNumerLokalu"] = i[3]
        to_import.loc[lista_adresatow.index(i), "AdresatKodPocztowy"] = i[4]
        to_import.loc[lista_adresatow.index(i), "AdresatMiejscowosc"] = i[5]
        to_import.loc[lista_adresatow.index(i), "AdresatKraj"] = i[6]
        to_import.loc[lista_adresatow.index(i), "Format"] = "S"
        to_import.loc[lista_adresatow.index(i), "KategoriaLubGwarancjaTerminu"] = "E"

    with pd.ExcelWriter(filepath, mode='w', engine='xlwt') as writer:
            to_import.to_excel(writer, sheet_name=sheets[0], index=False)




adresaci=[]
doc=DocxTemplate('templatka.docx')
file= pd.ExcelFile('dane.xlsx')
dane_adresowe={"miejscowosc":"","kod pocztowy":" ","ulica":"", "numer domu":"", "numer lokalu":" "}
dzisiejsza_data=datetime.date.today()
dzisiejsza_data=dzisiejsza_data.strftime("%d.%m.%Y")
nazwy_arkuszy=file.sheet_names
df=pd.read_excel('dane.xlsx')
i=0
wb = Workbook("dane_kopia.xlsx")
wb.save("dane_kopia.xlsx")
wb.close()
df=df.reset_index()

ile=int(input("Ile pism chcesz wygenerować?"))
for arkusz in nazwy_arkuszy:
    df=pd.read_excel('dane.xlsx',arkusz)
    for wiersz in df.itertuples():
        if i<ile:
            if pd.isna(wiersz.Dłużnik) or pd.isna(wiersz._3) or pd.isna(wiersz._4) or pd.isna(wiersz._5) or pd.isna(wiersz._6) or pd.isna(wiersz._15) or pd.isna(wiersz._14):
                continue
            if wiersz.Wygenerowano=='tak' and wiersz.Wygenerowano_pismo=="nie":
                i+=1
                imie,nazwisko=imie_i_nazwisko(wiersz.Dłużnik)
                oplata1=float(wiersz._14.split()[0])
                oplata2=float(wiersz._15.split()[0])
                kd_pocztowy=kod_pocztowy(str(wiersz._3))
                miejscowosc=(wiersz._4).capitalize()
                ulica=(wiersz._5).capitalize()
                numer_domu=wiersz._6
                id=wiersz._1
                if not(pd.isna(wiersz._7)):
                    numer_lokalu=wiersz._7
                else:
                    numer_lokalu='-'
                wniosek_dane={"data":dzisiejsza_data,"imie":imie, "nazwisko":nazwisko, "oplata1": "{:.2f}".format(oplata1), "oplata2":"{:.2f}".format(oplata2),"suma_oplat":"{:.2f}".format(oplata1+oplata2)}
                doc.render(wniosek_dane)
                adresaci.append(( str(str(nazwisko)+' '+str(imie)),ulica,numer_domu,numer_lokalu,kd_pocztowy,miejscowosc,'Polska'))
                if os.path.exists('pisma_sądowe')==False:
                    os.mkdir("pisma_sądowe",0o666)
                    print("utworzony folder pisma_sadowe")
                if os.path.exists("pisma_sądowe/"+str(dzisiejsza_data))==False:
                    os.mkdir(("pisma_sądowe/"+str(dzisiejsza_data)),0o666)
                    print("utoworzono folder z dzisiewierszsza datą")

                nazwa_pliku=(str(arkusz)+"_"+str(id)).upper()+'.docx'
                doc.save(nazwa_pliku)
                convert(nazwa_pliku,"pisma_sądowe/"+str(dzisiejsza_data))
                os.remove(nazwa_pliku)
                df.loc[wiersz.Index,"Wygenerowano_pismo"]="tak"

    with pd.ExcelWriter("dane_kopia.xlsx", mode='a',engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=arkusz, index=False)

workbook = openpyxl.load_workbook("dane_kopia.xlsx")
pusty = workbook['Sheet']
workbook.remove(pusty)
workbook.save("dane_kopia.xlsx")
file.close()
os.remove("dane.xlsx")
os.rename("dane_kopia.xlsx","dane.xlsx")

print(adresaci)
utworzXlsDoImportu(adresaci)