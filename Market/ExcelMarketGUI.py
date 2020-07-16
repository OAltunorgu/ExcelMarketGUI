import tkinter as tk
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
from xlwings.constants import InsertShiftDirection
import os
import datetime

class StokIslemleri(object):
    def __init__(self):
        self.Stok_Screen=tk.Toplevel(root)
        self.Stok_Screen.title("STOK İSLEM EKRANI")
        self.Stok_Screen.geometry("378x200")
        self.Stok_Screen.transient(root)
        
        
        
        
        
        self.StokUpdate=tk.Button(self.Stok_Screen,text="Ürün Güncelleme", width=25, height=5,bg="yellow",relief="ridge",command=self.UrunGuncelleme)
   
        self.StokAdd=tk.Button(self.Stok_Screen,text="Yeni Ürün Ekle", width=25, height=5, bg="yellow",relief="ridge",command=self.NewAdd) 
 
        self.StokUpdate.grid(row=0,column=0,padx=2,pady=3,sticky="w")

        self.StokAdd.grid(row=0,column=1,padx=2,pady=3,sticky="w")
        
        self.StokDelete=tk.Button(self.Stok_Screen,text="Ürün Silme", width=25, height=5,bg="red",relief="ridge",command=self.Delete)    
        self.StokReport=tk.Button(self.Stok_Screen,text="Raporlar", width=25, height=5,bg="lightblue",relief="ridge", command=self.Report)   
        self.StokDelete.grid(row=1,column=0,pady=2,sticky="w") 
        self.StokReport.grid(row=1,column=1,pady=2,sticky="w")
        
        
    def UrunGuncelleme(self):
        self.Stok_Screen=tk.Toplevel(root)
        self.Stok_Screen.title("ÜRÜN GÜNCELLEME EKRANI")
        self.Stok_Screen.geometry("600x500")
        self.Stok_Screen.transient(root)
        self.UrunBilgi_Frame=tk.Frame(self.Stok_Screen)
        self.UrunBilgi_Frame.grid(row=0,column=0)
        self.Stok_Liste=tk.Frame(self.Stok_Screen)
        self.Stok_Liste.grid(row=0,column=1)
        self.UrunKodu=tk.Label(self.UrunBilgi_Frame, text="Lütfen Ürün Kodu Giriniz")
        self.UrunKodu.grid(row=2,column=0)
        self.UrunKodu_Var=tk.IntVar()
        self.UrunKodu_Giris=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",\
                                     textvariable=self.UrunKodu_Var)
        self.UrunKodu_Giris.grid(row=2,column=1)
        
                

        
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Ürün Adı")
        self.UrunLabel.grid(row=5,column=0)
        self.UrunAdi_Var=tk.StringVar()
        self.UrunAdi=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunAdi_Var)
        self.UrunAdi.grid(row=5,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Raf Adet")
        self.UrunLabel.grid(row=6,column=0)
        self.UrunRafAdet_Var=tk.IntVar()
        self.UrunRafAdet=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunRafAdet_Var)
        self.UrunRafAdet.grid(row=6,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Stok Adet")
        self.UrunLabel.grid(row=7,column=0)
        self.UrunStokAdet_Var=tk.IntVar()
        self.UrunStokAdet=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunStokAdet_Var)
        self.UrunStokAdet.grid(row=7,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Ürün Geliş Fiyat")
        self.UrunLabel.grid(row=8,column=0)
        self.UrunGelis_Var=tk.IntVar()
        self.UrunGelis=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunGelis_Var)
        self.UrunGelis.grid(row=8,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="KDV") 
        self.UrunLabel.grid(row=9,column=0)
        self.UrunKDV_Var=tk.IntVar()
        self.UrunKDV=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKDV_Var)
        self.UrunKDV.grid(row=9,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Ürün Kar Oranı(%)") 
        self.UrunLabel.grid(row=10,column=0)
        self.UrunKarOranı_Var=tk.IntVar()
        self.UrunKarOranı=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKarOranı_Var)
        self.UrunKarOranı.grid(row=10,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="KDV Miktarı") 
        self.UrunLabel.grid(row=11,column=0)
        self.UrunKDVFiyat_Var=tk.IntVar()
        self.UrunKDVFiyat=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKDVFiyat_Var)
        self.UrunKDVFiyat.grid(row=11,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Kar Fiyat") 
        self.UrunLabel.grid(row=12,column=0)
        self.UrunKarFiyat_Var=tk.IntVar()
        self.UrunKarFiyat=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKarFiyat_Var)
        self.UrunKarFiyat.grid(row=12,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Satıs Fiyatı") 
        self.UrunLabel.grid(row=13,column=0)
        self.UrunSatisFiyat_Var=tk.IntVar()
        self.UrunSatisFiyat=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunSatisFiyat_Var)
        self.UrunSatisFiyat.grid(row=13,column=1)
        
        self.UrunEkle=tk.Button(self.UrunBilgi_Frame,text="Güncelle",bg="pink", command=self.UrunGuncelle)
        self.UrunEkle.grid(row=14,column=1)
        
        self.UrunEkle=tk.Button(self.UrunBilgi_Frame,text="Oku",bg="pink", command=self.UrunKoduOkuYaz)
        self.UrunEkle.grid(row=2,column=2)
        
        self.StokUrunlerListesi=[]
        

        
    def UrunKoduOkuYaz(self):
        UrunStokKodu=self.UrunKodu_Var.get()
        # Urun kodunu STOKLAR sheet'inden bulup detaylarını yazdır
        sheet=wb.sheets["STOKLAR"]
        LastCell = sheet.range(1,1).end('down').row # Number of rows
        
        Stok_Kontrol=False
        for row in range(2,LastCell+1):                
            if int(sheet.cells(row,1).value)==UrunStokKodu:
                    Stok_Kontrol=True
                    self.UrunAdi.insert("end",sheet.cells(row,2).value)
                    self.UrunRafAdet.insert("end",sheet.cells(row,3).value)
                    
                    self.UrunStokAdet.insert("end",sheet.cells(row,4).value)
                    
                    self.UrunGelis.insert("end",sheet.cells(row,5).value)
                    
                    self.UrunKDV.insert("end",sheet.cells(row,6).value)
                    
                    self.UrunKarOranı.insert("end",sheet.cells(row,7).value)
                    
                    self.UrunKDVFiyat.insert("end",sheet.cells(row,8).value)
                                   
                    self.UrunKarFiyat.insert("end",sheet.cells(row,9).value)
                    self.UrunSatisFiyat.insert("end",str(sheet.cells(row,10).value))
                    break
        if Stok_Kontrol==False:
                self.UrunAdi.insert("end","ÜRÜN STOKTA YOKTUR")
                self.NewAdd()
                
    def UrunGuncelle(self):
        self.sheet=wb.sheets["STOKLAR"]
        
        
        self.sheet.range("A5").api.Insert(InsertShiftDirection.xlShiftDown)
        
        
        wb.save()
        ##########################################################################################################
    
    def Delete(self):
        self.NewAdd_Screen=tk.Toplevel(root)
        self.NewAdd_Screen.title("YENİ ÜRÜN EKLEME")
        self.NewAdd_Screen.geometry("600x500")
        self.NewAdd_Screen.transient(root)       
        self.UrunBilgi_Frame=tk.Frame(self.NewAdd_Screen)
        self.UrunBilgi_Frame.grid(row=0,column=0)
        self.Stok_Liste=tk.Frame(self.NewAdd_Screen)
        self.Stok_Liste.grid(row=0,column=1)
        
        self.UrunKodu=tk.Label(self.UrunBilgi_Frame, text="Lütfen Ürün Kodu Giriniz")
        self.UrunKodu.grid(row=2,column=0)
        self.UrunKodu_Var=tk.IntVar()
        self.UrunKodu_Giris=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKodu_Var)
        self.UrunKodu_Giris.grid(row=2,column=1)

        self.UrunEkle=tk.Button(self.UrunBilgi_Frame,text="Güncelle",bg="pink", command=self.RowDelete)
        self.UrunEkle.grid(row=9,column=1)
    
    def Report(self):
        self.Stok_Screen=tk.Toplevel(root)
        self.Stok_Screen.title("RAPOR İŞLEMLERİ EKRANI")
        self.Stok_Screen.geometry("500x250")
        self.Stok_Screen.transient(root)
        self.UrunBilgi_Frame=tk.Frame(self.Stok_Screen)
        self.UrunBilgi_Frame.grid(row=0,column=0)
        self.Satis_Liste=tk.Frame(self.Stok_Screen)
        self.Satis_Liste.grid(row=0,column=1)
        self.UrunKodu=tk.Label(self.UrunBilgi_Frame, text="Kar İçin Ürün Kodu Giriniz")
        self.UrunKodu.grid(row=2,column=0)
        self.UrunAdi_Var=tk.IntVar()
        self.UrunKodu_Giris=tk.Entry(self.UrunBilgi_Frame,width=30,textvariable=self.UrunAdi_Var)
        self.UrunKodu_Giris.grid(row=2,column=1)
                
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Urun Kodu ve Kar:")
        self.UrunLabel.grid(row=3,column=0)
        self.UrunDetayi=tk.Text(self.UrunBilgi_Frame,width=22,height=10)
        self.UrunDetayi.grid(row=3,column=1)
        
        self.ToplamSatis_Var=tk.DoubleVar()
        self.TutarLabel=tk.Label(self.UrunBilgi_Frame,text="Toplam Kar:")
        self.TutarLabel.grid(row=4,column=0)
        self.TutarYaz=tk.Entry(self.UrunBilgi_Frame,width=30,textvariable=self.ToplamSatis_Var)
        self.TutarYaz.grid(row=4,column=1)
        self.ToplamSatis_Var.set(0.0)
        
        self.UrunKoduOku=tk.Button(self.UrunBilgi_Frame,text="Oku",bg="pink",command=self.UrunKoduOkuYaz1)
        self.UrunKoduOku.grid(row=2,column=2)

        
    def RowDelete(self):
        pass
    
    
    
    
    
    
    
    
    
    def NewAdd(self):
        self.NewAdd_Screen=tk.Toplevel(root)
        self.NewAdd_Screen.title("YENİ ÜRÜN EKLEME")
        self.NewAdd_Screen.geometry("600x500")
        self.NewAdd_Screen.transient(root)       
        self.UrunBilgi_Frame=tk.Frame(self.NewAdd_Screen)
        self.UrunBilgi_Frame.grid(row=0,column=0)
        self.Stok_Liste=tk.Frame(self.NewAdd_Screen)
        self.Stok_Liste.grid(row=0,column=1)
        
        self.UrunKodu=tk.Label(self.UrunBilgi_Frame, text="Lütfen Ürün Kodu Giriniz")
        self.UrunKodu.grid(row=2,column=0)
        self.UrunKodu_Var=tk.IntVar()
        self.UrunKodu_Giris=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKodu_Var)
        self.UrunKodu_Giris.grid(row=2,column=1)
                
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Ürün Adı")
        self.UrunLabel.grid(row=3,column=0)
        self.UrunAdi_Var=tk.StringVar()
        self.UrunDetayi=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunAdi_Var)
        self.UrunDetayi.grid(row=3,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Raf Adet")
        self.UrunLabel.grid(row=4,column=0)
        self.UrunRafAdet_Var=tk.IntVar()
        self.UrunRafAdet=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunRafAdet_Var)
        self.UrunRafAdet.grid(row=4,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Stok Adet")
        self.UrunLabel.grid(row=5,column=0)
        self.UrunStokAdet_Var=tk.IntVar()
        self.UrunStokAdet=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunStokAdet_Var)
        self.UrunStokAdet.grid(row=5,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Ürün Geliş Fiyat")
        self.UrunLabel.grid(row=6,column=0)
        self.UrunGelis_Var=tk.IntVar()
        self.UrunGelis=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunGelis_Var)
        self.UrunGelis.grid(row=6,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="KDV") 
        self.UrunLabel.grid(row=7,column=0)
        self.UrunKDV_Var=tk.IntVar()
        self.UrunKDV=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKDV_Var)
        self.UrunKDV.grid(row=7,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Ürün Kar Oranı(%)") 
        self.UrunLabel.grid(row=8,column=0)
        self.UrunKarOranı_Var=tk.IntVar()
        self.UrunKarOranı=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKarOranı_Var)
        self.UrunKarOranı.grid(row=8,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="KDV Miktarı") 
        self.UrunLabel.grid(row=9,column=0)
        self.UrunKDVFiyat_Var=tk.IntVar()
        self.UrunKDVFiyat=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKDVFiyat_Var)
        self.UrunKDVFiyat.grid(row=9,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Kar Fiyat") 
        self.UrunLabel.grid(row=10,column=0)
        self.UrunKarFiyat_Var=tk.IntVar()
        self.UrunKarFiyat=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunKarFiyat_Var)
        self.UrunKarFiyat.grid(row=10,column=1)
        
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Satıs Fiyat") 
        self.UrunLabel.grid(row=11,column=0)
        self.UrunSatisFiyat_Var=tk.IntVar()
        self.UrunSatisFiyat=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",textvariable=self.UrunSatisFiyat_Var)
        self.UrunSatisFiyat.grid(row=11,column=1)
        
        self.UrunEkle=tk.Button(self.UrunBilgi_Frame,text="Ekle",bg="pink",command=self.UrunEkle)
        self.UrunEkle.grid(row=12,column=2)
      

        
    def UrunEkle(self):
        
        UrunStokKodu=self.UrunKodu_Var.get()
        self.sheet=wb.sheets["STOKLAR"]
        LastCell = sheet.range(1,1).end('down').row # Number of rows
        
        Stok_Kontrol=False
        for row in range(2,LastCell+1):                
            if int(sheet.cells(row,1).value)==UrunStokKodu:
                    Stok_Kontrol=True
                    self.UrunAdi.insert("end",sheet.cells(row,2).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    self.UrunRafAdet.insert("end",sheet.cells(row,3).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    
                    self.UrunStokAdet.insert("end",sheet.cells(row,4).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    
                    self.UrunGelis.insert("end",sheet.cells(row,5).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    
                    self.UrunKDV.insert("end",sheet.cells(row,6).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    
                    self.UrunKarOranı.insert("end",sheet.cells(row,7).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    
                    self.UrunKDVFiyat.insert("end",sheet.cells(row,8).value).api.Insert(InsertShiftDirection.xlShiftDown)
                                   
                    self.UrunKarFiyat.insert("end",sheet.cells(row,9).value).api.Insert(InsertShiftDirection.xlShiftDown)
                    self.UrunSatisFiyat.insert("end",str(sheet.cells(row,10).value)).api.Insert(InsertShiftDirection.xlShiftDown)
                    break
                

    def UrunKoduOkuYaz1(self):
            UrunStokKodu=self.UrunAdi_Var.get()
            # Urun kodunu STOKLAR sheet'inden bulup detaylarını yazdır
            self.sheet=wb.sheets["STOKLAR"]
            LastCell = self.sheet.range(1,1).end('down').row # Number of rows
            self.UrunDetayi.delete("1.0","end")
            for row in range(2,LastCell+1):                
                if int(self.sheet.cells(row,1).value)==UrunStokKodu: 
                    self.UrunDetayi.insert("end",int(self.sheet.cells(row,1).value))
                    self.UrunDetayi.insert("end","\n")   
                    self.UrunDetayi.insert("end","KAR (Birim):"+str(self.sheet.cells(row,9).value))
                    self.UrunDetayi.insert("end","\n")

                    Kar=self.ToplamSatis_Var.get()
                    self.ToplamSatis_Var.set(Kar+self.sheet.cells(row,9).value)
                    


class SatisIslemleri(object):
    def __init__(self):
        self.Satis_Screen=tk.Toplevel(root)
        self.Satis_Screen.title("SATIŞ İŞLEMLERİ EKRANI")
        self.Satis_Screen.geometry("1300x650")
        self.Satis_Screen.transient(root)
        self.UrunBilgi_Frame=tk.Frame(self.Satis_Screen)
        self.UrunBilgi_Frame.grid(row=0,column=0)
        self.Satis_Liste=tk.Frame(self.Satis_Screen)
        self.Satis_Liste.grid(row=0,column=1)
        self.UrunKodu=tk.Label(self.UrunBilgi_Frame, text="Lütfen Ürün Kodu Giriniz")
        self.UrunKodu.grid(row=2,column=0)
        self.UrunAdi_Var=tk.IntVar()
        self.UrunKodu_Giris=tk.Entry(self.UrunBilgi_Frame,width=25,bg="lightblue",\
                                     textvariable=self.UrunAdi_Var)
        self.UrunKodu_Giris.grid(row=2,column=1)
                
        self.UrunLabel=tk.Label(self.UrunBilgi_Frame,text="Urun Detayı:")
        self.UrunLabel.grid(row=3,column=0)
        self.UrunDetayi=tk.Text(self.UrunBilgi_Frame,width=30,height=10)
        self.UrunDetayi.grid(row=3,column=1)
        
        self.ToplamSatis_Var=tk.DoubleVar()
        self.TutarLabel=tk.Label(self.UrunBilgi_Frame,text="TOPLAM SATIŞ TUTARI:")
        self.TutarLabel.grid(row=4,column=0)
        self.TutarYaz=tk.Entry(self.UrunBilgi_Frame,width=14,bg="lightblue",\
                               font="helvetica 24 bold",textvariable=self.ToplamSatis_Var)
        self.TutarYaz.grid(row=4,column=1)
        self.ToplamSatis_Var.set(0.0)
        
        self.UrunKoduOku=tk.Button(self.UrunBilgi_Frame,text="Oku",bg="pink",\
                                  command=self.UrunKoduOkuYaz)
        self.UrunKoduOku.grid(row=2,column=2)
        
        self.SatisListeLabel=tk.Label(self.Satis_Liste,text="SATIŞ SEPETİNDEKİ ÜRÜNLER")
        self.SatisListeLabel.grid(row=1,column=0)
        self.SatisSepeti=tk.Text(self.Satis_Liste,width=70,height=30)
        self.SatisSepeti.grid(row=2,column=0)
        
        self.UrunIptal=tk.Button(self.UrunBilgi_Frame,text="Ürün İptal",bg="pink",\
                                  command=self.UrunIptal)
        self.UrunIptal.grid(row=5,column=2)
        
        self.SatisIptal=tk.Button(self.UrunBilgi_Frame,text="Satış İptal",bg="pink",\
                                  command=self.SatisIptal)
        self.SatisIptal.grid(row=6,column=2)
        
        self.SatisBitir=tk.Button(self.UrunBilgi_Frame,text="Satış Bitir",bg="pink",\
                                  command=self.SatisBitir)
        self.SatisBitir.grid(row=7,column=2)
        
        self.SatisUrunlerListesi=[]
        
    def UrunKoduOkuYaz(self):
        UrunStokKodu=self.UrunAdi_Var.get()
        # Urun kodunu STOKLAR sheet'inden bulup detaylarını yazdır
        self.sheet=wb.sheets["STOKLAR"]
        LastCell = self.sheet.range(1,1).end('down').row # Number of rows
        self.UrunDetayi.delete("1.0","end")
        for row in range(2,LastCell+1):                
            if int(self.sheet.cells(row,1).value)==UrunStokKodu: 
                self.UrunDetayi.insert("end",int(self.sheet.cells(row,1).value))
                self.UrunDetayi.insert("end","\n")   
                self.UrunDetayi.insert("end",self.sheet.cells(row,2).value)
                self.UrunDetayi.insert("end","\n")
                self.UrunDetayi.insert("end",self.sheet.cells(row,3).value)
                self.UrunDetayi.insert("end","\n")
                self.UrunDetayi.insert("end","KDV = "+str(self.sheet.cells(row,6).value))
                self.UrunDetayi.insert("end","\n") 
                self.UrunDetayi.insert("end","Fiyat = "+str(self.sheet.cells(row,10).value))
                self.UrunDetayi.insert("end","\n")
                Tutar=self.ToplamSatis_Var.get()
                self.ToplamSatis_Var.set(Tutar+self.sheet.cells(row,10).value)
                
                self.SatisUrunlerListesi.append([int(self.sheet.cells(row,1).value),
                                                 self.sheet.cells(row,2).value,
                                                 self.sheet.cells(row,3).value,
                                                 self.sheet.cells(row,8).value,
                                                 self.sheet.cells(row,10).value])
    
                self.SatisSepeti.insert("end",int(self.sheet.cells(row,1).value))
                self.SatisSepeti.insert("end"," ")
                self.SatisSepeti.insert("end",self.sheet.cells(row,2).value)
                self.SatisSepeti.insert("end"," ")
                self.SatisSepeti.insert("end","KDV:"+str(self.sheet.cells(row,6).value))
                self.SatisSepeti.insert("end"," ")
                self.SatisSepeti.insert("end",self.sheet.cells(row,10).value)
                self.SatisSepeti.insert("end","\n")
                
    def LastEmptyRow(self,col=1):
        """ Find the last row in the worksheet that contains data.
        col: The column in which to look for the last cell containing data.
        """
        row=2
        while self.sheet.cells(row,1).value != None:
            row=row+1

        return row

    def UrunIptal(self):
        UrunStokKodu=self.UrunAdi_Var.get()  
        self.sheet=wb.sheets["STOKLAR"]
        LastCell = self.LastEmptyRow("STOKLAR") # Number of rows
        self.UrunDetayi.delete("1.0","end")
        Stok_Kontrol=False
        for row in range(2,LastCell+1):                
            if int(self.sheet.cells(row,1).value)==UrunStokKodu: 
                Stok_Kontrol=True
                self.UrunDetayi.insert("end",int(self.sheet.cells(row,1).value))
                self.UrunDetayi.insert("end","\n")   
                self.UrunDetayi.insert("end",self.sheet.cells(row,2).value)
                self.UrunDetayi.insert("end","\n")
                self.UrunDetayi.insert("end",self.sheet.cells(row,3).value)
                self.UrunDetayi.insert("end","\n")
                self.UrunDetayi.insert("end","KDV = "+str(self.sheet.cells(row,6).value))
                self.UrunDetayi.insert("end","\n") 
                self.UrunDetayi.insert("end","Fiyat = "+str(self.sheet.cells(row,10).value))
                self.UrunDetayi.insert("end","\n")
                break
        if Stok_Kontrol==False:
            self.UrunDetayi.insert("end","ÜRÜN STOK KODU YANLIŞTIR")
        else:
            Satis_Kontrol=False
            for x in self.SatisUrunlerListesi:
                if x[0]== UrunStokKodu:
                    Satis_Kontrol=True
                    Tutar=self.ToplamSatis_Var.get()
                    self.ToplamSatis_Var.set(Tutar-x[4])
                    self.SatisUrunlerListesi.remove(x)
                    break
            if Satis_Kontrol==False:
                self.UrunDetayi.insert("end","BU ÜRÜN SATIŞ LİSTESİNDE YOKTUR")
            else:                
                self.SatisSepeti.delete("1.0","end")
                for x in self.SatisUrunlerListesi:
                    self.SatisSepeti.insert("end",int(x[0]))
                    self.SatisSepeti.insert("end"," ")
                    self.SatisSepeti.insert("end",x[1])
                    self.SatisSepeti.insert("end"," ")
                    self.SatisSepeti.insert("end","KDV:"+str(x[3]*100))
                    self.SatisSepeti.insert("end"," ")
                    self.SatisSepeti.insert("end",x[4])
                    self.SatisSepeti.insert("end","\n")
                            
    
    def SatisIptal(self):
        self.UrunAdi_Var.set(0)
        self.UrunDetayi.delete("1.0","end")
        self.ToplamSatis_Var.set(0)
        self.SatisSepeti.delete("1.0","end")
    
    def SatisBitir(self):
        self.sheet=wb.sheets["STOKLAR"]
        LastCell = self.LastEmptyRow("STOKLAR") # Number of rows        
        for x in self.SatisUrunlerListesi:
            UrunKodu=x[0]
            for row in range(2,LastCell+1):
                if self.sheet.cells(row,1).value==UrunKodu:
                    self.sheet.cells(row,3).value=self.sheet.cells(row,3).value-1
                    break
        
        self.sheet=wb.sheets["SATIS_VERILERI"]
        LastCell = self.LastEmptyRow("SATIS_VERILERI") # Number of rows              

        now_datetime = datetime.datetime.now() 
        for x in self.SatisUrunlerListesi:
            self.sheet.cells(LastCell,1).value=now_datetime.strftime("%D")
            self.sheet.cells(LastCell,2).value=now_datetime.strftime("%X")
            if LastCell==2:
                self.sheet.cells(LastCell,3).value=1
            else:
                self.sheet.cells(LastCell,3).value = \
                self.sheet.cells(LastCell-1,3).value+1
            self.sheet.cells(LastCell,4).value = Username
            self.sheet.cells(LastCell,5).value = x[0]
            self.sheet.cells(LastCell,6).value = x[1]
            self.sheet.cells(LastCell,7).value = x[3]
            self.sheet.cells(LastCell,8).value = x[4]
            LastCell += 1            
        
        wb.save()
        self.UrunAdi_Var.set(0)
        self.UrunDetayi.delete("1.0","end")
        self.ToplamSatis_Var.set(0)
        self.SatisSepeti.delete("1.0","end")
    
        
def Login_Successful_OpenWindow():
    Login_success_screen.destroy()
    Login_Screen.destroy()
    Account_Screen.destroy()
    if "Stock" in Account_Type:
        StokIslem=StokIslemleri()
    elif "Satis" in Account_Type:
        SatisIslem=SatisIslemleri()
    else:
        print("Hesap tip tanımı yanlış!")
    
def Delete_password_not_recognised():
    Password_not_recog_screen.destroy()
    
def Delete_user_not_found_screen():
    User_not_found_screen.destroy()    

def Login_Successful():
    global Login_success_screen
    Login_success_screen = tk.Toplevel(Login_Screen)
    Login_success_screen.title("Success")
    Login_success_screen.geometry("150x100")
    tk.Label(Login_success_screen, text="Login Success").pack()
    tk.Button(Login_success_screen, text="OK",command=Login_Successful_OpenWindow).pack()

def Password_not_recognised():
    global Password_not_recog_screen
    Password_not_recog_screen = tk.Toplevel(Login_Screen)
    Password_not_recog_screen.title("Failure")
    Password_not_recog_screen.geometry("150x100")
    tk.Label(Password_not_recog_screen, text="Invalid Password ").pack()
    tk.Button(Password_not_recog_screen, text="OK", command=Delete_password_not_recognised).pack()

def User_not_found():
    global User_not_found_screen
    User_not_found_screen = tk.Toplevel(Login_Screen)
    User_not_found_screen.title("Failure")
    User_not_found_screen.geometry("150x100")
    tk.Label(User_not_found_screen, text="User Not Found").pack()
    tk.Button(User_not_found_screen, text="OK", command=Delete_user_not_found_screen).pack()

def Login_Verify():
    global Access_Right
    global Account_Type
    global Username
    
    Username = Username_Verify.get()
    Password = Password_Verify.get()
    Username_login_entry.delete(0, "end")
    Password_login_entry.delete(0, "end")

    List_of_files = os.listdir()
    if Username+".txt" in List_of_files:
        File = open(Username+".txt", "r")
        Verify = File.read().splitlines()
        Access_Right=Verify[2]
        Account_Type=Verify[3]
        if "False" in Access_Right:
            pass
        elif Password in Verify[1]:
            Login_Successful()
        else:
            Password_not_recognised()
    else:
        User_not_found()


def Login():
    global Login_Screen
    Login_Screen = tk.Toplevel(Account_Screen)
    Login_Screen.title("Login")
    Login_Screen.geometry("300x250")
    tk.Label(Login_Screen, text="Please enter details below to login").pack()
    tk.Label(Login_Screen, text="").pack()

    global Username_Verify
    global Password_Verify
    
    Username_Verify = tk.StringVar()
    Password_Verify = tk.StringVar()

    global Username_login_entry
    global Password_login_entry

    tk.Label(Login_Screen, text="Username * ").pack()
    Username_login_entry = tk.Entry(Login_Screen, textvariable=Username_Verify)
    Username_login_entry.pack()
    tk.Label(Login_Screen, text="").pack()
    tk.Label(Login_Screen, text="Password * ").pack()
    Password_login_entry = tk.Entry(Login_Screen, textvariable=Password_Verify, show= '*')
    Password_login_entry.pack()
    tk.Label(Login_Screen, text="").pack()
    tk.Button(Login_Screen, text="Login", width=10, height=1, command = Login_Verify).pack()

def Register_User():
    Username_info = Username.get()
    Password_info = Password.get()
    AccountType_info=AccountType.get()

    file = open(Username_info+".txt", "w")
    file.write(Username_info + "\n")
    file.write(Password_info + "\n")
    file.write("True \n")
    file.write(AccountType_info)
    file.close()

    Username_entry.delete(0, "end")
    Password_entry.delete(0, "end")
    AccountType_entry.delete(0, "end")

    tk.Label(Register_screen, text="Registration Success", fg="green", \
          font=("calibri", 11)).pack()
    

def Register():
    global Register_screen
    Register_screen = tk.Toplevel(Account_Screen)
    Register_screen.title("Register")
    Register_screen.geometry("300x250")

    global Username
    global Password
    global AccountType
    
    global Username_entry
    global Password_entry
    global AccountType_entry
    
    Username = tk.StringVar()
    Password = tk.StringVar()
    AccountType= tk.StringVar()
    
    tk.Label(Register_screen, text="Please enter details below", bg="blue").pack()
    tk.Label(Register_screen, text="").pack()
    Username_label = tk.Label(Register_screen, text="Username * ")
    Username_label.pack()
    Username_entry = tk.Entry(Register_screen, textvariable=Username)
    Username_entry.pack()
    Password_label = tk.Label(Register_screen, text="Password * ")
    Password_label.pack()
    Password_entry = tk.Entry(Register_screen, textvariable=Password, show='*')
    Password_entry.pack()
    
    AccountType_label = tk.Label(Register_screen, text="Account Type * ")
    AccountType_label.pack()
    AccountType_entry = tk.Entry(Register_screen, textvariable=AccountType)
    AccountType_entry.pack()
    
    tk.Label(Register_screen, text="").pack()
    tk.Button(Register_screen, text="Register", width=10, height=1, bg="blue",\
              command=Register_User).pack()

def Account_Login():    
    global Account_Screen
    Account_Screen=tk.Toplevel(root)
    Account_Screen.geometry("300x250")
    Account_Screen.title("Account Login")
    Account_Screen.transient(root)
    
    tk.Label(Account_Screen,text="Select Your Choice", bg="blue", width="300", height="2", font=("Calibri", 13)).pack()
    tk.Label(Account_Screen,text="").pack()
    tk.Button(Account_Screen,text="Login", height="2", width="30", command = Login).pack()
    tk.Label(Account_Screen,text="").pack()
    tk.Button(Account_Screen, text="Register", height="2", width="30", command=Register).pack()
     
root = tk.Tk() 
root.configure(background='light green') 
root.title("Market Yönetim GUI") 
root.geometry("375x475")

StokIslem=tk.Label(root,text="TEMİZ MARKET STOK İŞLEMLERİ", width=25, height=15, \
                        bg="yellow",relief="ridge")    
SatisIslem=tk.Label(root,text="TEMİZ MARKET SATIŞ İŞLEMLERİ", width=25, height=15, \
                        bg="yellow",relief="ridge")
StokIslem.grid(row=0,column=0,padx=2,pady=3,sticky="w")
SatisIslem.grid(row=0,column=1,padx=2,pady=3,sticky="w")

StokLogin=tk.Button(root,text="STOK KULLANICI GİRİŞİ", width=25, height=15, \
                        bg="lightblue",relief="ridge",command=Account_Login)    
SatisLogin=tk.Button(root,text="SATIŞ KULLANICI GİRİŞİ", width=25, height=15, \
                        bg="lightblue",relief="ridge",command=Account_Login)   
StokLogin.grid(row=1,column=0,pady=2,sticky="w") 
SatisLogin.grid(row=1,column=1,pady=2,sticky="w")

wb = xw.Book('C:\BLGM-416 Calısma\Deneme-Calısmam.xlsx') 

root.mainloop() 
    
