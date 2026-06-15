import streamlit as st
import pandas as pd
import qrcode
import re
import os
import shutil
from PIL import Image

# Sayfa Genişliği ve Kurumsal Başlık Ayarları
st.set_page_config(
    page_title="TAV Akıllı QR Kod Üretici", 
    page_icon="📱", 
    layout="centered"
)

# Dosya isimleri için Türkçe karakter temizleme
def dosya_adi_temizle(metin):
    sozluk = str.maketrans("çğıöşüÇĞİÖŞÜ ", "cgiosucgiosu_")
    return metin.translate(sozluk)

# Excel hücrelerindeki boşluk veya 'nan' kalıntılarını temizleme
def temizle(metin):
    if not metin or str(metin).lower() == 'nan':
        return ""
    return str(metin).replace('_x0000_', '').strip()

# Esnek Sütun Yakalama Fonksiyonu (Regex ile Akıllı Eşleşme)
def sutun_bul(mevcut_sutunlar, alternatifler):
    for sutun in mevcut_sutunlar:
        sutun_temiz = sutun.lower().strip()
        for alt in alternatifler:
            if alt in sutun_temiz:
                return sutun
    return None

# Arayüz Tasarımı
st.markdown("""
    <div style='text-align: center; padding-bottom: 20px;'>
        <h1 style='color: #0b2545;'>📱 TAV Akıllı QR Kod Üretim Merkezi</h1>
        <p style='color: #5a6b7c; font-size: 16px;'>
            Sütun isimlerinden ve yazım hatalarından bağımsız, akıllı vCard QR kod üretim paneli.
        </p>
    </div>
    <hr style='margin-top: 0; margin-bottom: 25px;'>
""", unsafe_allow_html=True)

# Dosya Yükleme Alanı
yuklenen_dosya = st.file_uploader(
    "Lütfen personel listesini içeren Excel (.xlsx) veya CSV dosyasını sürükleyip bırakın:", 
    type=["xlsx", "csv"]
)

if yuklenen_dosya is not None:
    try:
        if yuklenen_dosya.name.endswith('.csv'):
            df = pd.read_csv(yuklenen_dosya)
        else:
            df = pd.read_excel(yuklenen_dosya)
        
        # Sütun isimlerini normalize et
        orijinal_sutunlar = [str(col).replace('_x0000_', '').strip() for col in df.columns]
        df.columns = orijinal_sutunlar
        
        # Akıllı Sütun Eşleştirmeleri
        isim_sutunu = sutun_bul(orijinal_sutunlar, ['adı soyadı', 'ad soyad', 'isim', 'name', 'personel', 'ad', 'soyad'])
        unvan_sutunu = sutun_bul(orijinal_sutunlar, ['ünvan', 'unvan', 'title', 'görev', 'gorev'])
        sirket_sutunu = sutun_bul(orijinal_sutunlar, ['şirket', 'sirket', 'org', 'company', 'kurum'])
        email_sutunu = sutun_bul(orijinal_sutunlar, ['mail', 'e-mail', 'eposta', 'e-posta'])
        adres_sutunu = sutun_bul(orijinal_sutunlar, ['adres', 'address', 'lokasyon'])
        telefon_sutunu = sutun_bul(orijinal_sutunlar, ['iletişim', 'iletisim', 'telefon', 'tel', 'gsm', 'cep', 'phone'])

        if not i̇sim_sutunu:
            st.error("❌ İsim Sütunu Algılanamadı: Dosyanızda ad-soyad içeren sütunu otomatik bulamadık. Lütfen sütun başlığını kontrol edin.")
        else:
            st.success(f"📋 Dosya algılandı! İsimler **'{isim_sutunu}'** sütunundan okunacak. Toplam **{len(df)}** satır veri var.")
            
            # Üretim Butonu
            if st.button("🚀 Akıllı QR Kodları Toplu Üret", use_container_width=True):
                
                klasor_adi = "TAV_QR_Kodlari"
                if os.path.exists(klasor_adi):
                    shutil.rmtree(klasor_adi)
                os.makedirs(klasor_adi, exist_ok=True)
                
                ilerleme_cubugu = st.progress(0)
                durum_mesaji = st.empty()
                
                toplam_satir = len(df)
                uretilen_adet = 0
                gecerli_qr_listesi = [] # Önizleme için hafızada tutacağız

                for index, row in df.iterrows():
                    yuzde = int(((index + 1) / toplam_satir) * 100)
                    ilerleme_cubugu.progress(yuzde)
                    
                    tam_isim = temizle(row.get(isim_sutunu, ''))
                    if not tam_isim:
                        continue
                    
                    durum_mesaji.text(f"⏳ İşleniyor: {tam_isim}")
                    
                    # İsim/Soyisim Akıllı Ayrıştırma
                    isim_parcalari = [p for p in tam_isim.split(' ') if p]
                    if len(isim_parcalari) >= 1:
                        soyad = isim_parcalari[-1]
                        ad = " ".join(isim_parcalari[:-1]) if len(isim_parcalari) > 1 else ""
                    else:
                        continue

                    # Dinamik veri çekimi
                    unvan = temizle(row.get(unvan_sutunu, '')) if unvan_sutunu else ""
                    sirket = temizle(row.get(sirket_sutunu, '')) if sirket_sutunu else ""
                    email = temizle(row.get(email_sutunu, '')) if email_sutunu else ""
                    
                    adres = temizle(row.get(adres_sutunu, '')).replace('\n', ' ') if adres_sutunu else ""
                    adres = re.sub(r'\s+', ' ', adres)
                    
                    # Telefon Temizleme Algoritması
                    gsm_no = None
                    if telefon_sutunu:
                        iletisim_metni = str(row.get(telefon_sutunu, '')).strip()
                        if iletisim_metni and iletisim_metni.lower() != 'nan':
                            satirlar = iletisim_metni.split('\n')
                            for satir in satirlar:
                                if any(k in satir.upper() for k in ['GSM', 'CEP']):
                                    no_bul = re.search(r'([+\d][\d\s()\-.]+)', satir)
                                    if no_bul: gsm_no = no_bul.group(1).strip(); break
                            if not gsm_no:
                                for satir in satirlar:
                                    if any(k in satir.upper() for k in ['TEL', 'TELEFON']):
                                        no_bul = re.search(r'([+\d][\d\s()\-.]+)', satir)
                                        if no_bul: gsm_no = no_bul.group(1).strip(); break
                            if not gsm_no:
                                no_bul = re.search(r'([+\d][\d\s()\-.]+)', iletisim_metni)
                                if no_bul: gsm_no = no_bul.group(1).strip()

                    # vCard Oluşturma
                    vcard_satirlari = [
                        "BEGIN:VCARD",
                        "VERSION:3.0",
                        f"N;CHARSET=UTF-8:{soyad};{ad};;;",
                        f"FN;CHARSET=UTF-8:{tam_isim}",
                    ]

                    if sirket: vcard_satirlari.append(f"ORG;CHARSET=UTF-8:{sirket}")
                    if unvan: vcard_satirlari.append(f"TITLE;CHARSET=UTF-8:{unvan}")
                    if gsm_no: vcard_satirlari.append(f"TEL;TYPE=CELL:{gsm_no}")
                    if email:
                        vcard_satirlari.append(f"item1.EMAIL;TYPE=INTERNET:{email}")
                        vcard_satirlari.append("item1.X-ABLabel:E-posta")
                    if adres:
                        vcard_satirlari.append(f"item2.ADR;CHARSET=UTF-8:;;{adres};;;;")
                        vcard_satirlari.append("item2.X-ABLabel:Adres")

                    vcard_satirlari.append("END:VCARD")
                    vcard = "\r\n".join(vcard_satirlari)

                    # QR Oluşturma
                    qr = qrcode.QRCode(box_size=8, border=4)
                    qr.add_data(vcard.encode('utf-8'))
                    qr.make(fit=True)
                    img = qr.make_image(fill_color="black", back_color="white")

                    temiz_ad = dosya_adi_temizle(ad)
                    temiz_soyad = dosya_adi_temizle(soyad)
                    dosya_yolu = f"{klasor_adi}/{temiz_ad}_{temiz_soyad}.png"
                    img.save(dosya_yolu)
                    
                    # Önizleme listesine ekle
                    gecerli_qr_listesi.append({
                        "isim": tam_isim,
                        "unvan": unvan if unvan else "Personel",
                        "dosya": dosya_yolu
                    })
                    uretilen_adet += 1

                durum_mesaji.text("📦 Paket oluşturuluyor...")
                shutil.make_archive(klasor_adi, 'zip', klasor_adi)
                
                with open(f"{klasor_adi}.zip", "rb") as f:
                    zip_verisi = f.read()
                
                # Temizlik
                shutil.rmtree(klasor_adi)
                os.remove(f"{klasor_adi}.zip")
                
                durum_mesaji.empty()
                st.balloons()
                st.success(f"🎉 Toplam {uretilen_adet} QR Kod başarıyla hazırlandı!")
                
                # İNDİRME BUTONU
                st.download_button(
                    label="📥 Tüm QR Kodlarını Toplu İndir (ZIP)",
                    data=zip_verisi,
                    file_name="TAV_QR_Kodlari.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
                # CANLI ÖNİZLEME GALERİSİ (İstediğin yeni görsel alan)
                st.markdown("### 🔍 Üretilen QR Kodların Önizlemesi")
                st.caption("Aşağıda sistem tarafından yeni oluşturulan kartların önizlemesini görebilirsiniz:")
                
                # QR kodları şık bir ızgara (grid) halinde yan yana 3'erli gösterelim
                cols = st.columns(3)
                for idx, item in enumerate(gecerli_qr_listesi):
                    col_secim = cols[idx % 3]
                    with col_secim:
                        # Görseli açıp arayüze basıyoruz
                        # Klasör silindiği için hafızadaki nesne üzerinden veya geçici listeden basıyoruz
                        st.image(item["dosya"], use_container_width=True)
                        st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0b2545;'>{item['isim']}</div>", unsafe_allow_html=True)
                        st.markdown(f"<div style='text-align: center; font-size: 12px; color: #7a8b9c; margin-bottom: 20px;'>{item['unvan']}</div>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"⚠️ Sistemde bir sorun oluştu: {e}")
