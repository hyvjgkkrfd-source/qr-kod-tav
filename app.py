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
            Telefon Önizleme Modüllü Gelişmiş vCard Üretim Paneli
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
        
        # Akıllı Sütun Eşleştirmeleri (Tüm isimler 'orijinal_sutunlar' olarak eşitlendi)
        isim_sutunu = sutun_bul(orijinal_sutunlar, ['adı soyadı', 'ad soyad', 'isim', 'name', 'personel', 'ad', 'soyad'])
        unvan_sutunu = sutun_bul(orijinal_sutunlar, ['ünvan', 'unvan', 'title', 'görev', 'gorev'])
        sirket_sutunu = sutun_bul(orijinal_sutunlar, ['şirket', 'sirket', 'org', 'company', 'kurum'])
        email_sutunu = sutun_bul(orijinal_sutunlar, ['mail', 'e-mail', 'eposta', 'e-posta'])
        adres_sutunu = sutun_bul(orijinal_sutunlar, ['adres', 'address', 'lokasyon'])
        telefon_sutunu = sutun_bul(orijinal_sutunlar, ['iletişim', 'iletisim', 'telefon', 'tel', 'gsm', 'cep', 'phone'])

        if not isim_sutunu:
            st.error("❌ Dosyanızda ad-soyad içeren sütun otomatik bulamadık. Lütfen sütun başlığını kontrol edin.")
        else:
            st.success(f"📋 Dosya algılandı! Toplam **{len(df)}** satır veri var.")
            
            if st.button("🚀 Akıllı QR Kodları ve Telefon Önizlemelerini Üret", use_container_width=True):
                
                klasor_adi = "TAV_QR_Kodlari"
                if os.path.exists(klasor_adi):
                    try:
                        shutil.rmtree(klasor_adi)
                    except:
                        pass
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
                    
                    # Telefon Temizleme
                    gsm_no = ""
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
                    
                    # Bilgileri sakla
                    gecerli_qr_listesi.append({
                        "isim": tam_isim,
                        "unvan": unvan if unvan else "Personel",
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
                
                # CANLI TELEFON ÖNİZLEME ALANI
                st.markdown("### 🔍 Dijital Kartvizit Canlı Önizleme Paneli")
                st.caption("QR kod okutulduğunda telefonda belirecek kurumsal kart yapısı:")
                
                for item in gecerli_qr_listesi:
                    col1, col2 = st.columns([1, 1.3])
                    
                    with col1:
                        st.markdown("<br><br>", unsafe_allow_html=True)
                        if os.path.exists(item["dosya"]):
                            st.image(item["dosya"], width=230)
                        st.markdown(f"<p style='text-align: center; font-weight: bold; color: #0b2545;'>{item['isim']}</p>", unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div style="background-color: #f4f6f9; border: 12px solid #222; border-radius: 30px; padding: 20px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto; max-width: 320px; box-shadow: 0px 4px 15px rgba(0,0,0,0.1); margin: 10px auto;">
                            <div style="width: 70px; height: 70px; background-color: #0b2545; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 24px; font-weight: bold; margin: 0 auto 10px auto;">
                                {item['isim'][0]}
                            </div>
                            <div style="text-align: center; font-size: 18px; font-weight: bold; color: #111; margin-bottom: 2px;">{item['isim']}</div>
                            <div style="text-align: center; font-size: 13px; color: #0076ff; font-weight: 500; margin-bottom: 15px;">{item['unvan']}</div>
                            
                            <hr style="border: 0; border-top: 1px solid #dcdcdc; margin: 10px 0;">
                            
                            <div style="margin-bottom: 10px;">
                                <span style="font-size: 10px; color: #777; display: block; text-transform: uppercase;">Şirket</span>
                                <span style="font-size: 13px; color: #222; font-weight: 500;">{item['sirket']}</span>
                            </div>
                            <div style="margin-bottom: 10px;">
                                <span style="font-size: 10px; color: #777; display: block; text-transform: uppercase;">Cep Telefonu</span>
                                <span style="font-size: 13px; color: #0076ff; font-weight: 500;">{item['telefon']}</span>
                            </div>
                            <div style="margin-bottom: 10px;">
                                <span style="font-size: 10px; color: #777; display: block; text-transform: uppercase;">E-posta</span>
                                <span style="font-size: 13px; color: #222;">{item['email']}</span>
                            </div>
                            <div style="margin-bottom: 5px;">
                                <span style="font-size: 10px; color: #777; display: block; text-transform: uppercase;">Adres</span>
                                <span style="font-size: 12px; color: #333; line-height: 1.3;">{item['adres']}</span>
                            </div>
                            
                            <div style="background-color: #34c759; color: white; text-align: center; padding: 8px; border-radius: 10px; font-size: 12px; font-weight: bold; margin-top: 15px; cursor: default;">
                                ✓ Rehbere Kaydedilmeye Hazır
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    st.markdown("<hr style='border: 1px dashed #e2e8f0; margin: 30px 0;'>", unsafe_allow_html=True)
                
                # Temizlik işlemleri
                if os.path.exists(f"{klasor_adi}.zip"):
                    os.remove(f"{klasor_adi}.zip")

    except Exception as e:
        st.error(f"⚠️ Sistemde bir sorun oluştu: {e}")
