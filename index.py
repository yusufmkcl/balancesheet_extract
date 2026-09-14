import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
import os

# --- 1. SABİTLER VE YARDIMCI FONKSİYONLAR ---
SAYI_DESENI = re.compile(
    r'^\(?-?\d{1,3}(\.\d{3})+(,\d+)?\)?%?$'
    r'|^\(?-?\d+,\d+\)?%?$'
    r'|^-$'
)

DIPNOT_BELIRTEC_DESENI = re.compile(r'^(\(\*+\)|\(\d+\)|\*+)$')


def kelime_tipi(text):
    if SAYI_DESENI.match(text):
        return "SAYI"
    return "METIN"


def sayisal_sutun_sinirlarini_bul(gecerli_kelimeler, sayfa_bbox, bosluk_esigi=3):
    sayisal_kelimeler = [k for k in gecerli_kelimeler if kelime_tipi(k['text']) == 'SAYI']
    if not sayisal_kelimeler:
        return None

    sayisal_kelimeler.sort(key=lambda k: k['x1'])
    kumeler = [[sayisal_kelimeler[0]]]
    for kelime in sayisal_kelimeler[1:]:
        mevcut_kume = kumeler[-1]
        onceki_x1_max = max(k['x1'] for k in mevcut_kume)
        if kelime['x1'] - onceki_x1_max > bosluk_esigi:
            kumeler.append([kelime])
        else:
            mevcut_kume.append(kelime)

    sinirlar = [sayfa_bbox[0]]
    ilk_kume_min_x0 = min(k['x0'] for k in kumeler[0])
    sinirlar.append(ilk_kume_min_x0 - 2)

    for i in range(len(kumeler) - 1):
        bu_kumenin_sag_ucu = max(k['x1'] for k in kumeler[i])
        sonraki_kumenin_sol_ucu = min(k['x0'] for k in kumeler[i + 1])
        sinirlar.append((bu_kumenin_sag_ucu + sonraki_kumenin_sol_ucu) / 2)

    son_kume_max_x1 = max(k['x1'] for k in kumeler[-1])
    sinirlar.append(son_kume_max_x1 + 5)
    return sinirlar


def sayfalari_ayikla(girdi_metni):
    sayfalar = set()
    parcalar = girdi_metni.split(',')

    for parca in parcalar:
        parca = parca.strip()
        if not parca:
            continue

        if '-' in parca:
            try:
                baslangic, bitis = map(int, parca.split('-'))
                sayfalar.update(range(min(baslangic, bitis), max(baslangic, bitis) + 1))
            except ValueError:
                pass
        else:
            try:
                sayfalar.add(int(parca))
            except ValueError:
                pass

    return sorted(list(sayfalar))


# --- 2. ANA ÇIKARTMA FONKSİYONU ---
def extract_pdf_tables(pdf_yolu, excel_yolu, cekilecek_sayfalar):
    basarili_sayfa_sayisi = 0

    with pd.ExcelWriter(excel_yolu, engine='openpyxl') as writer:
        with pdfplumber.open(pdf_yolu) as pdf:
            for sayfa_no in cekilecek_sayfalar:
                st.info(f"⏳ [{sayfa_no}. Sayfa] İşleniyor...")

                try:
                    sayfa = pdf.pages[sayfa_no - 1]
                except IndexError:
                    st.error(f"❌ Hata: Belgede {sayfa_no}. sayfa bulunmuyor.")
                    continue

                tum_kelimeler = sayfa.extract_words(keep_blank_chars=False, extra_attrs=['fontname', 'size'])
                if not tum_kelimeler:
                    continue

                tum_kelimeler.sort(key=lambda w: w['top'])
                satirlar = []
                mevcut_satir = [tum_kelimeler[0]]
                satir_top = tum_kelimeler[0]['top']

                for w in tum_kelimeler[1:]:
                    if abs(w['top'] - satir_top) > 4:
                        satirlar.append(mevcut_satir)
                        mevcut_satir = [w]
                        satir_top = w['top']
                    else:
                        mevcut_satir.append(w)
                satirlar.append(mevcut_satir)

                satir_bloklari = []
                for satir in satirlar:
                    satir.sort(key=lambda w: w['x0'])
                    bloklar = []

                    ilk_tip = kelime_tipi(satir[0]['text'])
                    mevcut_blok = {'x0': satir[0]['x0'], 'x1': satir[0]['x1'], 'tip': ilk_tip, 'text': satir[0]['text']}

                    for w in satir[1:]:
                        suanki_tip = kelime_tipi(w['text'])
                        bosluk = w['x0'] - mevcut_blok['x1']

                        if bosluk > 15 or suanki_tip != mevcut_blok['tip']:
                            bloklar.append(mevcut_blok)
                            mevcut_blok = {'x0': w['x0'], 'x1': w['x1'], 'tip': suanki_tip, 'text': w['text']}
                        else:
                            mevcut_blok['x1'] = w['x1']
                            mevcut_blok['text'] += " " + w['text']

                    bloklar.append(mevcut_blok)
                    satir_bloklari.append({'satir': satir, 'bloklar': bloklar})

                sayfa_genisligi = sayfa.width
                max_izin_verilen_duvar = sayfa_genisligi * 0.75
                aciklama_x1_listesi = []

                for sb in satir_bloklari:
                    sayi_bloklari = [b for b in sb['bloklar'] if b['tip'] == 'SAYI']
                    if sayi_bloklari:
                        ilk_blok = sb['bloklar'][0]
                        if ilk_blok['tip'] == 'METIN' and ilk_blok['x1'] < max_izin_verilen_duvar:
                            aciklama_x1_listesi.append(ilk_blok['x1'])

                if aciklama_x1_listesi:
                    gorunmez_duvar = max(aciklama_x1_listesi) + 5
                else:
                    gorunmez_duvar = max_izin_verilen_duvar

                gecerli_satirlar = []
                gecerli_kelimeler = []
                onceki_satir_ihlal_mi = False
                onceki_satir_bottom = 0
                onceki_satir_son_kelime = None

                for sb in satir_bloklari:
                    ihlal_var = False
                    sadece_metin_mi = all(blok['tip'] == 'METIN' for blok in sb['bloklar'])

                    for blok in sb['bloklar']:
                        if blok['tip'] == 'METIN':
                            if blok['x0'] < (gorunmez_duvar - 5) and blok['x1'] > (gorunmez_duvar + 5):
                                ihlal_var = True
                                break

                    satir_top = min(w['top'] for w in sb['satir'])
                    if not ihlal_var and sadece_metin_mi and onceki_satir_ihlal_mi:
                        if (satir_top - onceki_satir_bottom) < 15:
                            kalkan_aktif = False
                            if len(sb['bloklar']) > 1:
                                kalkan_aktif = True

                            if not kalkan_aktif and onceki_satir_son_kelime and sb['satir']:
                                suanki_font = sb['satir'][0].get('fontname', '')
                                suanki_size = round(sb['satir'][0].get('size', 0), 1)
                                onceki_font = onceki_satir_son_kelime.get('fontname', '')
                                onceki_size = round(onceki_satir_son_kelime.get('size', 0), 1)

                                if suanki_font != onceki_font or abs(suanki_size - onceki_size) >= 0.5:
                                    kalkan_aktif = True

                            if kalkan_aktif:
                                ihlal_var = False
                            else:
                                ihlal_var = True

                    if not ihlal_var and sadece_metin_mi and sb['satir']:
                        ilk_kelime_metni = sb['satir'][0].get('text', '')
                        if DIPNOT_BELIRTEC_DESENI.match(ilk_kelime_metni):
                            ihlal_var = True

                    onceki_satir_ihlal_mi = ihlal_var
                    onceki_satir_bottom = max(w['bottom'] for w in sb['satir'])
                    if sb['satir']:
                        onceki_satir_son_kelime = sb['satir'][-1]
                    else:
                        onceki_satir_son_kelime = None

                    if not ihlal_var:
                        gecerli_satirlar.append(sb['satir'])
                        gecerli_kelimeler.extend(sb['satir'])

                dikey_sinirlar = sayisal_sutun_sinirlarini_bul(gecerli_kelimeler, sayfa.bbox, bosluk_esigi=3)

                if dikey_sinirlar is None:
                    st.warning(f"⚠️ Uyarı: {sayfa_no}. sayfada sayısal veri bulunamadı, extract_table ile deneniyor.")
                    tablo_verisi = sayfa.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text"})
                else:
                    tablo_verisi = []
                    for satir_kelimeleri in gecerli_satirlar:
                        satir_hucreleri = [""] * (len(dikey_sinirlar) - 1)

                        for w in satir_kelimeleri:
                            orta_nokta = (w['x0'] + w['x1']) / 2
                            for i in range(len(dikey_sinirlar) - 1):
                                if dikey_sinirlar[i] <= orta_nokta <= dikey_sinirlar[i + 1]:
                                    if satir_hucreleri[i]:
                                        satir_hucreleri[i] += " " + w['text']
                                    else:
                                        satir_hucreleri[i] = w['text']
                                    break

                        if any(hucre.strip() for hucre in satir_hucreleri):
                            tablo_verisi.append(satir_hucreleri)

                if tablo_verisi and len(tablo_verisi) > 1:
                    df = pd.DataFrame(tablo_verisi[1:], columns=tablo_verisi[0])
                    df = df.replace('\n', ' ', regex=True)

                    sekme_adi = f"Sayfa_{sayfa_no}"
                    df.to_excel(writer, sheet_name=sekme_adi, index=False)
                    st.success(f"✅ Başarılı! '{sekme_adi}' sekmesi dosyaya eklendi.")
                    basarili_sayfa_sayisi += 1
                else:
                    st.warning(f"⚠️ Uyarı: {sayfa_no}. sayfada işlenecek tablo bulunamadı.")

    return basarili_sayfa_sayisi


# --- 3. STREAMLIT WEB ARAYÜZÜ ---
st.set_page_config(page_title="PDF Tablo Çıkarıcı", page_icon="📊", layout="centered")

st.title("📊 PDF'den Excel'e Tablo Çıkarıcı")
st.markdown("Bir PDF dosyası yükleyin ve içindeki tabloları çekmek istediğiniz sayfaları belirtin. Sayısal veriler içeren bilanço raporlarına yönelik hazırlanmıştır, onun haricindeki tablolar düzgün işlenemeyebilir")

yuklenen_pdf = st.file_uploader("PDF Dosyasını Seçin", type=["pdf"])

sayfa_girdisi = st.text_input(
    "Çekilecek Sayfalar",
    placeholder="Örn: 16, 22, 60-63, 69",
    help="Tek sayfa için virgül, aralık belirtmek için tire kullanabilirsiniz."
)

if st.button("Tabloları Çıkart 🚀", type="primary"):

    cekilecek_sayfalar = sayfalari_ayikla(sayfa_girdisi)

    if yuklenen_pdf is None:
        st.error("Lütfen bir PDF dosyası yükleyin!")
    elif not cekilecek_sayfalar:
        st.error("Lütfen geçerli sayfa numaraları girin!")
    else:
        st.info(f"İşleme alınan sayfalar: {', '.join(map(str, cekilecek_sayfalar))}")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(yuklenen_pdf.read())
            pdf_gecici_yol = tmp_pdf.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            excel_gecici_yol = tmp_excel.name

        try:
            with st.spinner("Tablolar PDF'den ayrıştırılıyor, lütfen bekleyin..."):
                basarili_sayi = extract_pdf_tables(pdf_gecici_yol, excel_gecici_yol, cekilecek_sayfalar)

            if basarili_sayi > 0:
                st.balloons()
                st.success("🎉 İşlem tamam! Tüm tablolar Excel dosyasına kaydedildi.")

                with open(excel_gecici_yol, "rb") as f:
                    excel_verisi = f.read()

                st.download_button(
                    label="📥 Çıkarılan Tabloları İndir (Excel)",
                    data=excel_verisi,
                    file_name="Cikarilan_Tablolar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Girdiğiniz sayfalarda çıkartılabilecek hiçbir tablo bulunamadı.")

        except Exception as e:
            st.error(f"Beklenmeyen bir hata oluştu: {str(e)}")

        finally:
            if os.path.exists(pdf_gecici_yol):
                os.remove(pdf_gecici_yol)
            if os.path.exists(excel_gecici_yol):
                os.remove(excel_gecici_yol)
