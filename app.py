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

# Esnek Sütun Yakalama Fonksiyonu (Genişletilmiş Alternatifler)
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
            Sütun isimlerinden bağımsız, gelişmiş vCard QR kod üretim paneli.
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
        
        orijinal_sutunlar = [str(col).replace('_x0000_', '').strip() for col in df.columns]
        df.columns = orijinal_sutunlar
        
        # Akıllı Sütun Eşleştirmeleri (Telefon listesi maksimum seviyede genişletildi)
        isim_sutunu = sutun_bul(orijinal_sutunlar, ['adı soyadı', 'ad soyad', 'isim', 'name', 'personel', 'ad', 'soyad', 'kişi'])
        unvan_sutunu = sutun_bul(orijinal_sutunlar, ['ünvan', 'unvan', 'title', 'görev', 'gorev', 'pozisyon'])
        sirket_sutunu = sutun_bul(orijinal_sutunlar, ['şirket', 'sirket', 'org', 'company', 'kurum', 'firma'])
        email_sutunu = sutun_bul(orijinal_sutunlar, ['mail', 'e-mail', 'eposta', 'e-posta', 'email'])
        adres_sutunu = sutun_bul(orijinal_sutunlar, ['adres', 'address', 'lokasyon', 'yer'])
        telefon_sutunu = sutun_bul(orijinal_sutunlar, ['iletişim', 'iletisim', 'telefon', 'tel', 'gsm', 'cep', 'phone', 'no', 'numara', 'mob', 'irtibat'])

        if not isim_sutunu:
            st.error("❌ Dosyanızda ad-soyad içeren sütun otomatik bulamadık. Lütfen sütun başlığını kontrol edin.")
        else:
            st.success(f"📋 Dosya algılandı! Toplam **{len(df)}** satır veri var.")
            
            # Arka planda hangi sütunun ne olarak eşleştiğini kullanıcıya gösterelim (Debug kolaylığı için)
            with st.expander("🔍 Otomatik Eşleşen Sütun Başlıklarını Gör"):
                st.write(f"**İsim Sütunu:** {isim_sutunu}")
                st.write(f"**Telefon Sütunu:** {telefon_sutunu if telefon_sutunu else '⚠️ Bulunamadı'}")
                st.write(f"**E-posta Sütunu:** {email_sutunu if email_sutunu else '⚠️ Bulunamadı'}")
                st.write(f"**Unvan Sütunu:** {unvan_sutunu if unvan_sutunu else '⚠️ Bulunamadı'}")
            
            if st.button("🚀 Akıllı QR Kodları ve Önizlemeleri Üret", use_container_width=True):
                
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
                    if not tam_isim:
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
                    
                    # GELİŞMİŞ Telefon Yakalama Mantığı
                    gsm_no = ""
                    if telefon_sutunu:
                        ham_telefon = str(row.get(telefon_sutunu, '')).strip()
                        if ham_telefon and ham_telefon.lower() != 'nan':
                            # Önce hücredeki satırları tek tek kontrol et (GSM/CEP/TEL etiketleri için)
                            satirlar = ham_telefon.split('\n')
                            for satir in satirlar:
                                if any(k in satir.upper() for k in ['GSM', 'CEP', 'TEL', 'TELEFON', 'MOB']):
                                    no_bul = re.search(r'([+\d][\d\s()\-.]+)', satir)
                                    if no_bul: 
                                        gsm_no = no_bul.group(1).strip()
                                        break
                            
                            # Eğer etiketlerden hiçbir şey yakalanamadıysa, hücrenin içindeki ilk sayı dizisini çek
                            if not gsm_no:
                                no_bul = re.search(r'([+\d][\d\s()\-.]+)', ham_telefon)
                                if no_bul:
                                    gsm_no = no_bul.group(1).strip()
                                else:
                                    # Hücrede parantez/artı yoksa düz sayıları temizle al
                                    düz_rakamlar = "".join(re.findall(r'\d+', ham_telefon))
                                    if düz_rakamlar:
                                        gsm_no = düz_rakamlar

                    # vCard Blokları
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
                st.success(f"🎉 Toplam {uretilen_adet} QR Kod başarıyla hazırlandı!")
                
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
        st.error(f"⚠️ Sistemde bir sorun oluştu: {e}")
