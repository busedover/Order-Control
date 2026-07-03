import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from io import BytesIO

# Sayfa ayarları
st.set_page_config(page_title="CPD Allocation Revision & Audit Tool", layout="wide")

st.title("🛡 CPD Alokasyon Revizyon ve Audit Dashboard'u")
st.subheader("Tüm Müşteriler İçin On-Top Revizyon İhtiyaçları & Anlık Stok Denetim Analizi")

# --- DOSYA YÜKLEME ALANI (Sadece 3 Dosya) ---
st.sidebar.header("📂 Denetim Excel Dosyalarını Yükleyin")
orders_file = st.sidebar.file_uploader("1. Gelen Bekleyen Siparişler Excel'i", type=["xlsx", "xls"])
alloc_file = st.sidebar.file_uploader("2. Mevcut Alokasyon (Allocation) Raporu", type=["xlsx", "xls"])
stock_file = st.sidebar.file_uploader("3. Güncel Stok Excel'i", type=["xlsx", "xls"])

# Katalog dosyasını arka planda otomatik köprü olarak okuyoruz
fiyat_dosya_adi = "fiyatlar 2026.xlsx"

@st.cache_data
def katalog_koprusunu_yukle():
    if os.path.exists(fiyat_dosya_adi):
        df_fiyat_raw = pd.read_excel(fiyat_dosya_adi, engine="openpyxl")
        df_fiyat_raw.columns = df_fiyat_raw.columns.str.strip()
        
        fiyat_barkod_col = "EAN Cod-UM"
        fiyat_material_col = "Material"
        
        if fiyat_barkod_col in df_fiyat_raw.columns and fiyat_material_col in df_fiyat_raw.columns:
            # Barkod temizleme
            df_fiyat_raw[fiyat_barkod_col] = df_fiyat_raw[fiyat_barkod_col].astype(str).str.strip()
            df_fiyat_raw[fiyat_barkod_col] = df_fiyat_raw[fiyat_barkod_col].apply(lambda x: re.sub(r'^\.+|\.+$', '', x))
            df_fiyat_raw[fiyat_barkod_col] = df_fiyat_raw[fiyat_barkod_col].apply(lambda x: x.split('.')[0] if '.' in x else x)
            df_fiyat_raw[fiyat_barkod_col] = df_fiyat_raw[fiyat_barkod_col].apply(lambda x: re.sub(r'\D', '', x))
            
            # Material temizleme
            df_fiyat_raw[fiyat_material_col] = df_fiyat_raw[fiyat_material_col].astype(str).str.strip()
            df_fiyat_raw[fiyat_material_col] = df_fiyat_raw[fiyat_material_col].apply(lambda x: x.split('.')[0] if '.' in x else x)
            df_fiyat_raw[fiyat_material_col] = df_fiyat_raw[fiyat_material_col].str.lstrip('0')
            
            df_clean = df_fiyat_raw[[fiyat_barkod_col, fiyat_material_col]].dropna().drop_duplicates(subset=[fiyat_barkod_col])
            df_clean.rename(columns={fiyat_barkod_col: "Barkod"}, inplace=True)
            return df_clean
    return None

df_catalog_bridge = katalog_koprusunu_yukle()

if df_catalog_bridge is not None:
    st.sidebar.success("✅ Barkod-Material Köprüsü otomatik bağlandı!")
else:
    st.sidebar.error(f"❌ '{fiyat_dosya_adi}' dosyası bulunamadı! Köprü çalışmayabilir.")

if orders_file and alloc_file and stock_file and df_catalog_bridge is not None:
    # Verileri oku
    df_orders = pd.read_excel(orders_file, engine="openpyxl")
    df_alloc = pd.read_excel(alloc_file, engine="openpyxl")
    df_stock = pd.read_excel(stock_file, engine="openpyxl")
    
    # Kolon başlıklarındaki boşlukları temizleme
    df_orders.columns = df_orders.columns.str.strip()
    df_alloc.columns = df_alloc.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    
    st.success("✅ Tüm analiz dosyaları başarıyla yüklendi!")
    
    # Kritik Sütun Tanımları
    siparis_barkod_col = "Barkod"
    stok_material_col = "Material"
    stok_net_avail_col = "Net avail."
    
    # Alokasyon dosyasındaki tahsis adetini belirten kolonu bulma
    alokasyon_adet_col = None
    for col in ["Tahsis Edilen Adet", "Tahsis Miktarı", "Alokasyon Adeti", "Tahsis Adedi", "Allocated Qty", "Allocated", "Alokasyon Miktarı"]:
        if col in df_alloc.columns:
            alokasyon_adet_col = col
            break
            
    kalan_adet_col = "Kalan adet" if "Kalan adet" in df_alloc.columns else ("Kalan Adet" if "Kalan Adet" in df_alloc.columns else None)
    
    # Kolon Kontrolleri
    orders_ok = siparis_barkod_col in df_orders.columns and "Sipariş Miktarı" in df_orders.columns
    alloc_ok = siparis_barkod_col in df_alloc.columns and alokasyon_adet_col is not None and kalan_adet_col is not None
    stock_ok = stok_material_col in df_stock.columns and stok_net_avail_col in df_stock.columns
    
    if orders_ok and alloc_ok and stock_ok:
        
        # --- BARKOD & MALZEME TEMİZLEME ---
        def barkod_temizle(seri):
            s_str = seri.astype(str).str.strip()
            s_str = s_str.apply(lambda x: re.sub(r'^\.+|\.+$', '', x))
            s_str = s_str.apply(lambda x: x.split('.')[0] if '.' in x else x)
            s_str = s_str.apply(lambda x: re.sub(r'\D', '', x))
            return s_str

        def malzeme_temizle(seri):
            s_str = seri.astype(str).str.strip()
            s_str = s_str.apply(lambda x: x.split('.')[0] if '.' in x else x)
            s_str = s_str.str.lstrip('0')
            return s_str

        df_orders[siparis_barkod_col] = barkod_temizle(df_orders[siparis_barkod_col])
        df_alloc[siparis_barkod_col] = barkod_temizle(df_alloc[siparis_barkod_col])
        df_stock[stok_material_col] = malzeme_temizle(df_stock[stok_material_col])
        df_stock[stok_net_avail_col] = pd.to_numeric(df_stock[stok_net_avail_col], errors='coerce').fillna(0)

        # --- BARKOD BAZLI STOK HESAPLAMA ---
        df_stock_grouped = df_stock.groupby(stok_material_col)[stok_net_avail_col].sum().reset_index()
        df_merged_stock = pd.merge(df_catalog_bridge, df_stock_grouped, on="Material", how="inner")
        df_barcode_stock = df_merged_stock.groupby("Barkod")[stok_net_avail_col].sum().reset_index()

        # --- AUDIT (DENETİM) BİRLEŞTİRME ---
        df_orders_sub = df_orders[[siparis_barkod_col, "Müşteri Adı", "Sipariş Miktarı"]]
        df_alloc_sub = df_alloc[[siparis_barkod_col, "Müşteri Adı", alokasyon_adet_col, kalan_adet_col]]
        
        # Müşteri ve Barkod bazında eşleştirme (Tüm Müşteriler)
        df_audit = pd.merge(df_orders_sub, df_alloc_sub, on=[siparis_barkod_col, "Müşteri Adı"], how="outer")
        df_audit["Sipariş Miktarı"] = df_audit["Sipariş Miktarı"].fillna(0)
        df_audit[alokasyon_adet_col] = df_audit[alokasyon_adet_col].fillna(0)
        df_audit[kalan_adet_col] = df_audit[kalan_adet_col].fillna(0)
        
        # Depo stoğunu ekleyelim
        df_audit = pd.merge(df_audit, df_barcode_stock, left_on=siparis_barkod_col, right_on="Barkod", how="left")
        df_audit[stok_net_avail_col] = df_audit[stok_net_avail_col].fillna(0)
        df_audit.drop(columns=["Barkod_y"], errors="ignore", inplace=True)
        df_audit.rename(columns={"Barkod_x": "Barkod"}, errors="ignore", inplace=True)
        
        # Barkod bazında toplam kalan alokasyonu bulalım (Tüm Müşteri Siparişleri Toplamı)
        df_audit["Toplam Kalan Alokasyon"] = df_audit.groupby("Barkod")[kalan_adet_col].transform("sum")
        
        # Boştaki Stok Formülü: Net avail. - Toplam Kalan Alokasyon
        df_audit["Boştaki Stok"] = df_audit[stok_net_avail_col] - df_audit["Toplam Kalan Alokasyon"]
        df_audit["Boştaki Stok"] = df_audit["Boştaki Stok"].apply(lambda x: max(0, x))

        # --- AUDIT REVİZYON ALGORİTMASI ---
        def audit_karar_motoru(row):
            kalan = row[kalan_adet_col]
            boştaki_stok = row["Boştaki Stok"]
            
            # Kural: Kalan sipariş adedi var ve depoda sizin belirttiğiniz formüle göre boşta serbest stok kalmış!
            if boştaki_stok > 0 and kalan > 0:
                ekstra_verilebilir = min(kalan, boştaki_stok)
                if ekstra_verilebilir > 0:
                    return "💡 Boşta Stok Var! Alokasyonu Artır", f"Depoda serbest stok var. Bu müşteriye +{int(ekstra_verilebilir)} adet on-top ekstra tahsis açın."
                    
            return "✅ Uyumlu", "Mevcut tahsis güncel depo stoğuna ve kalan talebe tam uygundur."

        # Analizi çalıştıralım
        audit_sonucları = df_audit.apply(audit_karar_motoru, axis=1)
        df_audit["Denetim Durumu"] = [x[0] for x in audit_sonucları]
        df_audit["Önerilen Aksiyon Planı"] = [x[1] for x in audit_sonucları]

        # --- YAZARAK ARAMA ÖZELLİKLİ FİLTRELEME ALANI ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Audit Filtreleme")
        
        # "Distribütör" ifadesi yerine genel "Müşteri" ifadesi getirildi
        tum_musteriler = sorted(df_audit['Müşteri Adı'].dropna().unique().tolist())
        secilen_musteriler = st.sidebar.multiselect("🏢 Müşteri Seçin (Zincir, Distribütör vb.)", options=tum_musteriler, placeholder="Müşteri adı...")
        
        durum_tipleri = ["Tümü", "💡 Boşta Stok Var! Alokasyonu Artır", "✅ Uyumlu"]
        secilen_durum = st.sidebar.selectbox("📋 Denetim Durumuna Göre Süz", durum_tipleri)
        
        # Filtreleri uygulama
        df_filtered = df_audit.copy()
        if len(secilen_musteriler) > 0:
            df_filtered = df_filtered[df_filtered['Müşteri Adı'].isin(secilen_musteriler)]
        if secilen_durum != "Tümü":
            df_filtered = df_filtered[df_filtered['Denetim Durumu'] == secilen_durum]

        # --- AUDIT KPI KARTLARI (ADET BAZLI) ---
        total_order_qty = df_filtered["Sipariş Miktarı"].sum()
        total_alloc_qty = df_filtered[alokasyon_adet_col].sum()
        total_kalan_qty = df_filtered[kalan_adet_col].sum()
        
        ekstra_stok_count = len(df_filtered[df_filtered["Denetim Durumu"].str.contains("Boşta Stok")])
        
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Toplam Bekleyen Sipariş Adedi", f"{int(total_order_qty):,} Adet")
        kpi2.metric("Mevcut Alokasyon Adedi", f"{int(total_alloc_qty):,} Adet", delta=f"{int(total_kalan_qty):,} Kalan Adet", delta_color="inverse")
        
        # Dinamik Durum KPI Kartı
        if ekstra_stok_count > 0:
            kpi3.metric("💡 Ekstra Tahsis Açılabilir", f"{ekstra_stok_count} Satır", delta="Boşta Stok Fırsatı!")
        else:
            kpi3.metric("✅ Audit Durumu", "Tam Uyumlu", delta="Tüm Alokasyonlar Güvende!")

        # --- TABLO GÖSTERİMİ ---
        st.markdown("---")
        st.subheader("📋 Alokasyon Revizyon Audit Raporu")
        st.caption("💡 İpucu: Listede süzülen herhangi bir hücredeki barkod veya aksiyon planını çift tıklayarak anında kopyalayabilirsiniz (Ctrl+C).")
        
        display_cols = [
            "Müşteri Adı", "Barkod", "Sipariş Miktarı", alokasyon_adet_col, kalan_adet_col,
            stok_net_avail_col, "Boştaki Stok", "Denetim Durumu", "Önerilen Aksiyon Planı"
        ]
        
        def color_audit_durumu(val):
            if "Boşta Stok" in val:
                return 'color: #1E90FF; font-weight: bold;'
            else:
                return 'color: green; font-weight: bold;'

        st.dataframe(
            df_filtered[display_cols].style.format({
                'Sipariş Miktarı': '{:,.0f}',
                alokasyon_adet_col: '{:,.0f}',
                kalan_adet_col: '{:,.0f}',
                stok_net_avail_col: '{:,.0f}',
                'Boştaki Stok': '{:,.0f}'
            }).map(color_audit_durumu, subset=['Denetim Durumu']),
            use_container_width=True
        )

        # --- EXCEL OLARAK DIŞARI AKTARMA ---
        st.markdown("### 📥 Revizyon ve Audit Raporunu Excel Olarak İndir")
        
        def to_excel(df_export):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Alokasyon Audit Raporu')
            processed_data = output.getvalue()
            return processed_data

        excel_data = to_excel(df_filtered[display_cols])
        st.download_button(
            label="📊 Audit Raporunu İndir (.xlsx)",
            data=excel_data,
            file_name="cpd_alokasyon_audit_raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    else:
        st.warning("⚠ Yüklenen dosyaların kolon isimlerini kontrol edin:")
        if not orders_ok:
            st.error("❌ Sipariş dosyasında 'Barkod' ve 'Sipariş Miktarı' olmalı.")
        if not alloc_ok:
            st.error(f"❌ Alokasyon dosyasında 'Barkod', tahsis miktarı ve '{kalan_adet_col}' kolonu olmalı.")
        if not stock_ok:
            st.error(f"❌ Stok dosyasında 'Material' ve '{stok_net_avail_col}' olmalı.")
else:
    st.info("💡 Lütfen sol menüden 'Siparişler', 'Mevcut Alokasyon' ve 'Stok' dosyalarını yükleyin. Denetim analizleri arka planda anlık olarak tamamlanacaktır.")
