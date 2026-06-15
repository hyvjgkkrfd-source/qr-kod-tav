import streamlit as st
import pandas as pd
import qrcode
import re
import os
import shutil

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

# Arayüz Tasarımı
st.markdown("""
    <div style='text-align: center; padding-bottom: 20px;'>
        <h1 style='color: #0b2545;'>📱 TAV Akıllı QR Kod Üretim Merkezi</h1>
        <p style='color: #5a6b7c; font-size: 16px;'>
            TAV Kartvizit Formatına Özel Geliştirilmiş QR Kod Üretim Paneli
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
        # TAV şablonundaki üst boşlukları aşmak için dosyayı önce normal okuyoruz
        if yuklenen_dosya.name.endswith('.csv'):
            test_df = pd.read_csv(yuklenen_dosya, nrows=10)
        else:
            test_df = pd.read_excel(yuklenen_dosya, nrows=10)
        
        # Gerçek başlık satırını bulma (İçinde 'Adı Soyadı' geçen satırı arıyoruz)
        skip_rows_value = 0
        for i, row in test_df.iterrows():
            row_str = " ".join([str(val).lower() for val in row.values])
            if 'adı soyadı' in row_str or 'adı' in row_str:
                skip_rows_value = i + 1
                break
        
        # Dosyayı bulduğumuz doğru satırdan itibaren yeniden yüklüyoruz
        if yuklenen_dosya.name.endswith('.csv'):
            df = pd.read_csv(yuklenen_dosya, skiprows=skip_rows_value)
        else:
            df = pd.read_excel(yuklenen_dosya, skiprows=skip_rows_value)
            
        # Sütun isimlerini tamamen temizle
        df.columns = [str(col).replace('_x0000_', '').strip() for col in df.columns]
        
        # TAV Şablonu Sütun Doğrulaması
        isim_sutunu = 'Adı Soyadı' if 'Adı Soyadı' in df.columns else df.columns[0]
        unvan_sutunu = 'Ünvanı (Türkçe)' if 'Ünvanı (Türkçe)' in df.columns else None
        sirket_sutunu = 'Şirket Adı (Türkçe)' if 'Şirket Adı (Türkçe)' in df.columns else None
        email_sutunu = 'e-mail hesabı' if 'e-mail hesabı' in df.columns else None
        adres_sutunu = 'Adres (Türkçe)' if 'Adres (Türkçe)' in df.columns else None
        telefon_sutunu = 'İletişim bilgileri (Türkçe)' if 'İletişim bilgileri (Türkçe)' in df.columns else None

        st.success(f"📋 TAV Kartvizit Şablonu başarıyla çözümlendi! Toplam **{len(df)}** personel işlenecek.")
        
        if st.button("🚀 QR Kodları ve Önizlemeleri Üret", use_container_width=True):
            
            klasor_adi = "TAV_QR_Kodlari"
            if os.path.exists(klasor_adi):
                try: shutil.rmtree(klasor_adi)
                except: pass
            os.makedirs(klasor_adi, exist_ok=True)
            
            ilerleme_cubugu = st.progress(0)
            durum_mesaji = st.empty()
            
            toplam_satir = len(df)
            uretilen_adet = 0
            gecerli_qr_listesi = []

            for index, row in df.iterrows():
                yuzde = int(((index + 1) / toplam_satir) * 100)
                ilerleme_cubugu.progress(yuzde)
                
                tam_isim = temizle(row.get(isim_sutunu, ''))
                if not tam_isim or tam_isim.lower() == 'nan':
                    continue
                
                durum_mesaji.text(f"⏳ İşleniyor: {tam_isim}")
                
                isim_parcalari = [p for p in tam_isim.split(' ') if p]
                if len(isim_parcalari) >= 1:
                    soyad = isim_parcalari[-1]
                    ad = " ".join(isim_parcalari[:-1]) if len(isim_parcalari) > 1 else ""
                else:
                    continue

                unvan = temizle(row.get(unvan_sutunu, '')) if unvan_sutunu else ""
                sirket = temizle(row.get(sirket_sutunu, '')) if sirket_sutunu else ""
                email = temizle(row.get(email_sutunu, '')) if email_sutunu else ""
                adres = temizle(row.get(adres_sutunu, '')).replace('\n', ' ') if adres_sutunu else ""
                adres = re.sub(r'\s+', ' ', adres)
                
                # Telefon Ayıklama Algoritması
                gsm_no = ""
                if telefon_sutunu:
                    iletisim_metni = str(row.get(telefon_sutunu, '')).strip()
                    if iletisim_metni and iletisim_metni.lower() != 'nan':
                        satirlar = re.split(r'[\n\r]+', iletisim_metni)
                        
                        for satir in satirlar:
                            if any(k in satir.upper() for k in ['MOBİL', 'MOBIL', 'GSM', 'CEP']):
                                no_bul = re.search(r'([\d\s()\-.]+)', satir.split(':')[-1])
                                if no_bul:
                                    gsm_no = no_bul.group(1).strip()
                                    gsm_no = re.sub(r'^[^\d+]+', '', gsm_no)
                                    break
                        
                        if not gsm_no:
                            for satir in satirlar:
                                if any(k in satir.upper() for k in ['TEL', 'TELEFON', 'PHONE']):
                                    no_bul = re.search(r'([\d\s()\-.]+)', satir.split(':')[-1])
                                    if no_bul:
                                        gsm_no = no_bul.group(1).strip()
                                        gsm_no = re.sub(r'^[^\d+]+', '', gsm_no)
                                        break
                        
                        if not gsm_no:
                            no_bul = re.search(r'([+\d][\d\s()\-.]+)', iletisim_metni)
                            if no_bul:
                                gsm_no = no_bul.group(1).strip()

                # vCard Formatlama
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

                # QR Oluşturma (HATA VEREN KISIM DÜZELTİLDİ: qr.make(fit=True))
                qr = qrcode.QRCode(box_size=8, border=4)
                qr.add_data(vcard.encode('utf-8'))
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")

                temiz_ad = dosya_adi_temizle(ad)
                temiz_soyad = dosya_adi_temizle(soyad)
                dosya_yolu = f"{klasor_adi}/{temiz_ad}_{temiz_soyad}.png"
                img.save(dosya_yolu)
                
                gecerli_qr_listesi.append({
                    "isim": tam_isim,
                    "unvan": unvan if unvan else "Belirtilmedi",
                    "sirket": sirket if sirket else "TAV",
                    "telefon": gsm_no if gsm_no else "Belirtilmedi",
                    "email": email if email else "Belirtilmedi",
                    "adres": adres if adres else "Belirtilmedi",
                    "dosya": dosya_yolu
                })
                uretilen_adet += 1

            durum_mesaji.text("📦 Paket oluşturuluyor...")
            shutil.make_archive(klasor_adi, 'zip', klasor_adi)
            
            with open(f"{klasor_adi}.zip", "rb") as f:
                zip_verisi = f.read()
            
            durum_mesaji.empty()
            st.balloons()
            st.success(f"🎉 Harika! Toplam {uretilen_adet} personelin QR kodu başarıyla üretildi.")
            
            # İNDİRME BUTONU
            st.download_button(
                label="📥 Tüm QR Kodlarını Toplu İndir (ZIP)",
                data=zip_verisi,
                file_name="TAV_QR_Kodlari.zip",
                mime="application/zip",
                use_container_width=True
            )
            
            # ÖNİZLEME ALANI
            st.markdown("### 🔍 Üretilen Kartlar ve Rehber Önizlemeleri")
            
            for item in gecerli_qr_listesi:
                col1, col2 = st.columns([1, 1.2])
                
                with col1:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if os.path.exists(item["dosya"]):
                        st.image(item["dosya"], width=220)
                
                with col2:
                    st.subheader(f"👤 {item['isim']}")
                    st.caption(f"💼 {item['unvan']} | 🏛️ {item['sirket']}")
                    st.info(f"📞 **Cep Telefonu:** {item['telefon']}")
                    st.code(f"✉️ E-posta: {item['email']}\n📍 Adres: {item['adres']}", language="text")
                    st.success("✓ Telefon kamerasıyla okutulmaya hazır.")
                    
                st.markdown("---")
            
            if os.path.exists(f"{klasor_adi}.zip"):
                os.remove(f"{klasor_adi}.zip")

    except Exception as e:
        st.error(f"⚠️ Dosya işlenirken bir sorun oluştu: {e}")
