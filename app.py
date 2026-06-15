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
