    #TOPSÄ°S
import pandas as pd
import numpy as np

def topsis(excel_path):
    # Excel dosyasini oku ve DataFrame'e donustur
    df_excel = pd.read_excel(excel_path)
    maliyet = df_excel.iloc[1:2].copy()
    agirlik = df_excel.iloc[:1].copy()
    df = df_excel.iloc[2:]

    # Kac satir kac sutun?
    satir_sayisi = df.shape[0]
    sutun_sayisi = df.shape[1]

    # Kareler Toplamlari ve Karekok
    karelertoplamlari = []
    karekok = []
    for j in range(sutun_sayisi):
        karetop = 0
        for i in range(satir_sayisi):
            karetop = karetop + (df.iloc[i, j] ** 2)
        karelertoplamlari.append(karetop)
        karekok.append(karetop ** 0.5)

    # Normalize Islemi
    normalize_df = df.copy()  # df verilerini kopyala
    for i in range(sutun_sayisi):
        normalize_df.iloc[:, i] = df.iloc[:, i] / karekok[i]

    # Agirlikli Normalize İslemi
    anormalize_df = normalize_df.copy()  # normalize_df verilerini kopyala
    for i in range(sutun_sayisi):
        anormalize_df.iloc[:, i] = normalize_df.iloc[:, i] * agirlik.iloc[0, i]

    aplus = []
    aminus = []
    for i in range(sutun_sayisi):
        if maliyet.iloc[0, i] == "maliyet":
            aminus.append(anormalize_df.iloc[:, i].max())  # Maksimum degeri kullan
            aplus.append(anormalize_df.iloc[:, i].min())  # Minimum degeri kullan
        else:
            aminus.append(anormalize_df.iloc[:, i].min())  # Minimum degeri kullan
            aplus.append(anormalize_df.iloc[:, i].max())  # Maksimum degeri kullan

    siplus = []
    for j in range(satir_sayisi):
        siplushesap = 0
        for i in range(sutun_sayisi):
            siplushesap = siplushesap + (anormalize_df.iloc[j, i] - aplus[i]) ** 2
        siplus.append(siplushesap ** 0.5)

    siminus = []
    for j in range(satir_sayisi):
        siminushesap = 0
        for i in range(sutun_sayisi):
            siminushesap = siminushesap + (anormalize_df.iloc[j, i] - aminus[i]) ** 2
        siminus.append(siminushesap ** 0.5)

    #Siralama ve Ci
    ci = []
    cihesap = 0
    for i in range(satir_sayisi):
        cihesap = siminus[i] / (siminus[i] + siplus[i])
        ci.append(cihesap)
    for i in range(satir_sayisi):
        max_deger = max(ci)
        max_index = ci.index(max_deger)
        print(f"Deger: {max_deger}, Index: {max_index + 1}")
        ci[max_index] = -99999


excel_yolu = "C:/Users/gokha/OneDrive/Masaüstü/İstatistik/Tez/1.1-Topsis.xlsx"
topsis(excel_yolu)



