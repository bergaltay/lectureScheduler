import pulp
import pandas as pd
import numpy as np

# Dersler
#'DERS KODU': {
#        'SECTION KODU': [('GUN', BAŞLANGIÇ SAATİ, BİTİŞ SAATİ), ('GUN', BAŞLANGIÇ SAATİ, BİTİŞ SAATİ)],
#        '02': [('Salı', 11, 13), ('Perşembe', 11, 12)],
#        '03': [('Pazartesi', 11, 13), ('Salı', 11, 12)]
#    },
#
#
#
dersler = {
    'CMPE 223': {
        '02': [('Perşembe', 14, 17)]
    },
    'MATH 240': {
        '01': [('Salı', 16, 17), ('Çarşamba', 13, 15)],
        '02': [('Salı', 11, 13), ('Çarşamba', 11, 12)],
        '03': [('Salı', 11, 13), ('Çarşamba', 11, 12)]
    },
    'MATH 250': {
        '01': [('Pazartesi', 13, 15), ('Perşembe', 14, 15)],
        '02': [('Salı', 14, 16), ('Perşembe', 13, 14)],
        '03': [('Salı', 9, 11), ('Cuma', 14, 15)]
    },
    'LIBE 130': {
        '01': [('Pazartesi', 13, 16)],
        '02': [('Salı', 12, 15)],
        '03': [('Çarşamba', 15, 18)],
        '04': [('Perşembe', 9, 12)]
    },
    'BIO 110': {
        '01': [('Perşembe', 13, 16)],
        '02': [('Cuma', 9, 12)],
        '03': [('Cuma', 13, 16)]
    },
    'CMPE 371': {
        '03': [('Cuma', 9, 12)],
        '04': [('Cuma', 14, 17)],
        '05': [('Perşembe', 9, 12)]
    }
}


def dersProgramiKombisanyonlariOlustur():
    kombinasyonlar = []
    dersler_listesi = list(dersler.keys())
    siniflar = {ders: list(sinifler.keys()) for ders, sinifler in dersler.items()}

    all_combinations = np.array(np.meshgrid(*[siniflar[ders] for ders in dersler_listesi])).T.reshape(-1,len(dersler_listesi))

    for ders_comb in all_combinations:
        problem = pulp.LpProblem(f"Ders Programı Kombinasyonu", pulp.LpMinimize)

        secilen_siniflar = pulp.LpVariable.dicts("SecilenSinif",
                                                 [(ders, sinif) for ders in dersler for sinif in dersler[ders]],
                                                 cat='Binary')

        for ders, sinif in zip(dersler_listesi, ders_comb):
            problem += secilen_siniflar[(ders, sinif)] == 1

        gunler = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma']
        for gun in gunler:
            for saat in range(8, 19):
                problem += pulp.lpSum([secilen_siniflar[(ders, sinif)]
                                       for ders in dersler
                                       for sinif in dersler[ders]
                                       for (g, baslangic, bitis) in dersler[ders][sinif]
                                       if g == gun and baslangic <= saat < bitis]) <= 1

        # kontrol et
        try:
            problem.solve()
            if pulp.LpStatus[problem.status] == 'Optimal':
                current_program = {}
                for ders in dersler:
                    for sinif in dersler[ders]:
                        if pulp.value(secilen_siniflar[(ders, sinif)]) == 1:
                            current_program[(ders, sinif)] = 1
                kombinasyonlar.append(current_program)
        except Exception as e:
            print(f"Bir hata oluştu: {e}")

    return kombinasyonlar


def ders_programini_excel(program, sayfa_adi, writer):
    gunler = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma']
    saatler = np.arange(8, 19)  # 08:00 ile 19:00 saatleri arasında bak

    df = pd.DataFrame(index=saatler, columns=gunler)

    for ders in dersler:
        for sinif in dersler[ders]:
            if (ders, sinif) in program:
                for (gun, baslangic, bitis) in dersler[ders][sinif]:
                    gun_index = gunler.index(gun)
                    for saat in range(baslangic, bitis):
                        df.loc[saat, gunler[gun_index]] = f'{ders} {sinif}'

    df.to_excel(writer, sheet_name=sayfa_adi)


# Tüm kombinasyonları oluştur
kombinasyonlar = dersProgramiKombisanyonlariOlustur()

# Excel dosyasını oluştur
dosya_adi = 'dersProgramlari.xlsx'
with pd.ExcelWriter(dosya_adi, engine='openpyxl') as writer:
    # Her kombinasyonu Excel dosyasına yaz
    for i, program in enumerate(kombinasyonlar):
        sayfa_adi = f'Kombinasyon {i + 1}'
        ders_programini_excel(program, sayfa_adi, writer)
        print(f'Kombinasyon {i + 1} Excel dosyasına kaydedildi: Sayfa Adı: {sayfa_adi}')
