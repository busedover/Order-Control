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
catalog_file = st.sidebar.file_uploader("2. Katalog Raporu Excel'i (Köprü)", type=["xlsx", "xls"])
stock_file = st.sidebar.file_uploader("3. Güncel Stok Excel'i", type=["xlsx", "xls"])

if orders_file and catalog_file and stock_file:
    # Verileri oku
    df_orders = pd.read_excel(orders_file, engine="openpyxl")
    df_catalog = pd.read_excel(catalog_file, engine="openpyxl")
    df_stock = pd.read_excel(stock_file, engine="openpyxl")
    
    # Kolon başlıklarındaki boşlukları temizleme
    df_orders.columns = df_orders.columns.str.strip()
    df_catalog.columns = df_catalog.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    
    st.success("✅ Tüm analiz dosyaları başarıyla yüklendi!")
    
    # Kritik Sütun Tanımları
    siparis_barkod_col = "Barkod"
    katalog_material_col = "Material"
    katalog_ean_col = "EAN Cod-UM"
    stok_material_col = "Material"
    stok_net_avail_col = "Net avail."
    
    # Alokasyon dosyasındaki (1. Gelen Bekleyen Siparişler) tahsis ve kalan adet kolonlarını dinamik bulma
    alokasyon_adet_col = None
    for col in ["Tahsis Edilen Adet", "Tahsis Miktarı", "Alokasyon Adeti", "Tahsis Adedi", "Allocated Qty", "Allocated", "Alokasyon Miktarı"]:
        if col in df_orders.columns:
            alokasyon_adet_col = col
            break
            
    kalan_adet_col = "Kalan adet" if "Kalan adet" in df_orders.columns else ("Kalan Adet" if "Kalan Adet" in df_orders.columns else None)
    
    # Kolon Kontrolleri
    orders_ok = siparis_barkod_col in df_orders.columns and "Sipariş Miktarı" in df_orders.columns and alokasyon_adet_col is not None and kalan_adet_col is not None
    catalog_ok = katalog_material_col in df_catalog.columns and katalog_ean_col in df_catalog.columns
    stock_ok = stok_material_col in df_stock.columns and stok_net_avail_col in df_stock.columns
    
    if orders_ok and catalog_ok and stock_ok:
        
        # --- BARKOD & MALZEME TEMİZLEME MOTORLARI ---
        def gelismis_barkod_temizle(seri):
            s_str = seri.astype(str).str.strip()
            s_str = s_str.apply(lambda x: re.sub(r'^\.+|\.+$', '', x))
            def e_notation_duzelt(val):
                if 'e' in val.lower():
                    try:
                        return str(int(float(val)))
                    except:
                        pass
                return val
            s_str = s_str.apply(e_notation_duzelt)
            s_str = s_str.apply(lambda x: x.split('.')[0] if '.' in x else x)
            s_str = s_str.apply(lambda x: re.sub(r'\D', '', x))
            return s_str

        def malzeme_kodunu_temizle(seri):
            s_str = seri.astype(str).str.strip()
            s_str = s_str.apply(lambda x: x.split('.')[0] if '.' in x else x)
            s_str = s_str.str.lstrip('0')
            return s_str

        # Format temizliği
        df_orders[siparis_barkod_col] = gelismis_barkod_temizle(df_orders[siparis_barkod_col])
        df_catalog[katalog_material_col] = malzeme_kodunu_temizle(df_catalog[katalog_material_col])
        df_catalog[katalog_ean_col] = gelismis_barkod_temizle(df_catalog[katalog_ean_col])
        df_stock[stok_material_col] = malzeme_kodunu_temizle(df_stock[stok_material_col])
        df_stock[stok_net_avail_col] = pd.to_numeric(df_stock[stok_net_avail_col], errors='coerce').fillna(0)

        # --- BARKOD BAZLI STOK HESAPLAMA ---
        df_stock_grouped = df_stock.groupby(stok_material_col)[stok_net_avail_col].sum().reset_index()
        
        # Köprüyü dinamik olarak yüklediğiniz Katalog Excel'inden kuruyoruz!
        df_cat_bridge = df_catalog[[katalog_material_col, katalog_ean_col]].dropna().drop_duplicates()
        df_merged_stock = pd.merge(df_cat_bridge, df_stock_grouped, on=katalog_material_col, how="inner")
        
        # Barkod bazında toplam depo stoğu
        df_barcode_stock_sum = df_merged_stock.groupby(katalog_ean_col)[stok_net_avail_col].sum().reset_index()
        df_barcode_stock_sum.rename(columns={katalog_ean_col: "Barkod"}, inplace=True)

        # --- S SİPARİŞ VE STOK BİRLEŞTİRME ---
        df_final = pd.merge(df_orders, df_barcode_stock_sum, on="Barkod", how="left")
        df_final[stok_net_avail_col] = df_final[stok_net_avail_col].fillna(0)
        
        # --- BARKOD BAZLI TOPLAM KALAN ALOKASYON VE BOŞTAKİ STOK HESABI ---
        df_final["Toplam Kalan Alokasyon"] = df_final.groupby("Barkod")[kalan_adet_col].transform("sum")
        
        df_final["Boştaki Stok"] = df_final[stok_net_avail_col] - df_final["Toplam Kalan Alokasyon"]
        df_final["Boştaki Stok"] = df_final["Boştaki Stok"].apply(lambda x: max(0, x))

        # --- AUDIT REVİZYON ALGORİTMASI ---
        def audit_karar_motoru(row):
            kalan = row[kalan_adet_col]
            boştaki_stok = row["Boştaki Stok"]
            
            # Kural: Kalan sipariş adedi var ve depoda boşta serbest stok kalmış!
            if boştaki_stok > 0 and kalan > 0:
                ekstra_verilebilir = min(kalan, boştaki_stok)
                if ekstra_verilebilir > 0:
                    return "💡 Boşta Stok Var! Alokasyonu Artır", f"Depoda serbest stok var. Bu müşteriye +{int(ekstra_verilebilir)} adet on-top ekstra tahsis açın."
                    
            return "✅ Uyumlu", "Mevcut tahsis güncel depo stoğuna ve kalan talebe tam uygundur."

        # Analizi çalıştıralım
        audit_sonucları = df_final.apply(audit_karar_motoru, axis=1)
        df_final["Denetim Durumu"] = [x[0] for x in audit_sonucları]
        df_final["Önerilen Aksiyon Planı"] = [x[1] for x in audit_sonucları]

        # --- YAZARAK ARAMA ÖZELLİKLİ FİLTRELEME ALANI ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Audit Filtreleme")
        
        tum_musteriler = sorted(df_final['Müşteri Adı'].dropna().unique().tolist())
        secilen_musteriler = st.sidebar.multiselect("🏢 Müşteri Seçin (Zincir, Distribütör vb.)", options=tum_musteriler, placeholder="Müşteri adı...")
        
        durum_tipleri = ["Tümü", "💡 Boşta Stok Var! Alokasyonu Artır", "✅ Uyumlu"]
        secilen_durum = st.sidebar.selectbox("📋 Denetim Durumuna Göre Süz", durum_tipleri)
        
        # Filtreleri uygulama
        df_filtered = df_final.copy()
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
            st.error(f"❌ Sipariş dosyasında 'Barkod', 'Sipariş Miktarı', '{alokasyon_adet_col}' ve '{kalan_adet_col}' olmalı.")
        if not catalog_ok:
            st.error(f"❌ Katalog dosyasında '{katalog_material_col}' ve '{katalog_ean_col}' olmalı.")
        if not stock_ok:
            st.error(f"❌ Stok dosyasında 'Material' ve '{stok_net_avail_col}' olmalı.")
else:
    st.info("💡 Lütfen sol menüden 'Siparişler', 'Katalog' ve 'Stok' excel dosyalarını yükleyin. Denetim analizleri arka planda anlık olarak tamamlanacaktır.")
