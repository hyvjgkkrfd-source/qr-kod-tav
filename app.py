import streamlit as st
import pandas as pd
import qrcode
import re
import os
import shutil

# Sayfa Genişliği ve Kurumsal Başlık Ayarları
st.set_page_config(
    page_title="TAV QR Kod Üretici", 
    page_icon="📱", 
    layout="centered"
)

# Sadece dosya isimlerinde güvenliği sağlamak için Türkçe karakter temizleme
def dosya_adi_temizle(metin):
    sozluk = str.maketrans("çğıöşüÇĞİÖŞÜ ", "cgiosucgiosu_")
    return metin.translate(sozluk)

# Excel hücrelerindeki boşluk veya 'nan' kalıntılarını temizleme
def temizle(metin):
    if not metin or str(metin).lower() == 'nan':
        return ""
    return str(metin).replace('_x0000_', '').strip()

# Arayüz Tasarımı (HTML/CSS Dokunuşları ile)
st.markdown("""
    <div style='text-align: center; padding-bottom: 20px;'>
        <h1 style='color: #0b2545;'>📱 TAV QR Kod Üretim Merkezi</h1>
        <p style='color: #5a6b7c; font-size: 16px;'>
            Excel veya CSV formatındaki personel listenizi yükleyin, vCard uyumlu QR kodlarınızı anında toplu ZIP olarak indirin.
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
        # Dosya uzantısına göre veriyi içeri aktarma
        if yuklenen_dosya.name.endswith('.csv'):
            df = pd.read_csv(yuklenen_dosya)
        else:
            df = pd.read_excel(yuklenen_dosya)
        
        # Sütun isimlerindeki görünmez Excel kalıntılarını arındırma
        df.columns = [str(col).replace('_x0000_', '').strip() for col in df.columns]
        
        # Kritik Sütun Kontrolü (Hata yönetimini arayüze taşıyoruz)
        if 'Adı Soyadı' not in df.columns:
            st.error("❌ Dosya Yapısı Hatalı: Yüklediğiniz dosyada 'Adı Soyadı' sütunu bulunamadı! Lütfen Excel başlıklarınızı kontrol edin.")
        else:
            st.info(f"📋 Dosya başarıyla algılandı. Toplam **{len(df)}** satır veri işlenmeye hazır.")
            
            # Üretim Butonu
            if st.button("🚀 QR Kodları Toplu Üret", use_container_width=True):
                
                # Önceki üretim kalıntılarını temizleme ve temiz klasör açma
                klasor_adi = "TAV_QR_Kodlari"
                if os.path.exists(klasor_adi):
                    shutil.rmtree(klasor_adi)
                os.makedirs(klasor_adi, exist_ok=True)
                
                # Canlı İlerleme Çubuğu ve Durum Metinleri
                ilerleme_cubugu = st.progress(0)
                durum_mesaji = st.empty()
                
                toplam_satir = len(df)
                uretilen_adet = 0

                for index, row in df.iterrows():
                    # İlerleme yüzdesini güncelleme
                    yuzde = int(((index + 1) / toplam_satir) * 100)
                    ilerleme_cubugu.progress(yuzde)
                    
                    tam_isim = temizle(row.get('Adı Soyadı', ''))
                    if not tam_isim:
                        continue
                    
                    durum_mesaji.text(f"⏳ İşleniyor: {tam_isim}")
                    
                    # İsim ve Soyisim Ayrıştırma
                    isim_parcalari = tam_isim.split(' ')
                    soyad = isim_parcalari[-1]
                    ad = " ".join(isim_parcalari[:-1]) if len(isim_parcalari) > 1 else ""

                    unvan = temizle(row.get('Ünvanı (Türkçe)', ''))
                    sirket = temizle(row.get('Şirket Adı (Türkçe)', ''))
                    email = temizle(row.get('e-mail hesabı', ''))
                    
                    # Adres satırlarındaki kırılmaları temizleme
                    adres = temizle(row.get('Adres (Türkçe)', '')).replace('\n', ' ')
                    adres = re.sub(r'\s+', ' ', adres)
                    
                    web_sitesi = temizle(row.get('Web Sitesi', row.get('Web Sitesi (Türkçe)', '')))
                    iletisim_metni = str(row.get('İletişim bilgileri (Türkçe)', '')).strip()

                    # Akıllı Telefon Numarası Ayrıştırma Filtreleri
                    gsm_no = None
                    if iletisim_metni and iletisim_metni.lower() != 'nan':
                        satirlar = iletisim_metni.split('\n')
                        for satir in satirlar:
                            if any(k in satir.upper() for k in ['GSM', 'CEP']):
                                no_bul = re.search(r'([+\d][\d\s()\-.]+)', satir)
                                if no_bul:
                                    gsm_no = no_bul.group(1).strip()
                                    break
                        if not gsm_no:
                            for satir in satirlar:
                                if any(k in satir.upper() for k in ['TEL', 'TELEFON']):
                                    no_bul = re.search(r'([+\d][\d\s()\-.]+)', satir)
                                    if no_bul:
                                        gsm_no = no_bul.group(1).strip()
                                        break
                        if not gsm_no:
                            no_bul = re.search(r'([+\d][\d\s()\-.]+)', iletisim_metni)
                            if no_bul:
                                gsm_no = no_bul.group(1).strip()

                    # Uluslararası Standartlara Uygun vCard Blokları
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
                    if web_sitesi:
                        link = web_sitesi if web_sitesi.startswith("http") else "https://" + web_sitesi
                        vcard_satirlari.append(f"item3.URL:{link}")
                        vcard_satirlari.append("item3.X-ABLabel:Web Sitesi")

                    vcard_satirlari.append("END:VCARD")
                    vcard = "\r\n".join(vcard_satirlari)

                    # Kamera Optik Okuma Performansı İçin Optimize Edilmiş QR (box_size=8)
                    qr = qrcode.QRCode(box_size=8, border=4)
                    qr.add_data(vcard.encode('utf-8'))
                    qr.make(fit=True)
                    img = qr.make_image(fill_color="black", back_color="white")

                    # İşletim sistemi uyumluluğu için temizlenmiş dosya adı
                    temiz_ad = dosya_adi_temizle(ad)
                    temiz_soyad = dosya_adi_temizle(soyad)
                    dosya_adi = f"{klasor_adi}/{temiz_ad}_{temiz_soyad}.png"
                    
                    img.save(dosya_adi)
                    uretilen_adet += 1

                # Süreç Bitimi ve Arşivleme
                durum_mesaji.text("📦 QR Kodları paketleniyor, lütfen bekleyin...")
                shutil.make_archive(klasor_adi, 'zip', klasor_adi)
                
                # ZIP dosyasını arayüz indirme butonuna bağlamak için okuma
                with open(f"{klasor_adi}.zip", "rb") as f:
                    zip_verisi = f.read()
                
                # Temizlik (Sunucu hafızasında çöp bırakmamak için)
                shutil.rmtree(klasor_adi)
                os.remove(f"{klasor_adi}.zip")
                
                # Başarı Bildirimi ve Kutlama Efekti
                durum_mesaji.empty()
                st.balloons()
                st.success(f"🎉 Harika! Toplam {uretilen_adet} personelin QR kodu başarıyla üretildi.")
                
                # İndirme Butonu Aktif Ediliyor
                st.download_button(
                    label="📥 Tüm QR Kodlarını Toplu İndir (ZIP)",
                    data=zip_verisi,
                    file_name="TAV_QR_Kodlari.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"⚠️ Dosya okunurken beklenmeyen bir hata oluştu: {e}")
