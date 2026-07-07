import streamlit as st
import pandas as pd
import numpy as np
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
    # Verileri oku - Sipariş sayfası 'Sheet1' üzerinde çalışıyoruz
    df_orders = pd.read_excel(orders_file, sheet_name="Sheet1", engine="openpyxl")
    df_catalog = pd.read_excel(catalog_file, engine="openpyxl")
    df_stock = pd.read_excel(stock_file, engine="openpyxl")
    
    # Kolon başlıklarındaki boşlukları temizleme
    df_orders.columns = df_orders.columns.str.strip()
    df_catalog.columns = df_catalog.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    
    st.success("✅ Tüm analiz dosyaları başarıyla yüklendi!")
    
    # --- KOLON BAŞLIKLARINI EŞLEŞTİRME ---
    siparis_material_col = "Material"
    katalog_material_col = "Material"
    katalog_ean_col = "EAN Cod-UM"
    stok_material_col = "Material"
    stok_net_avail_col = "Net avail."
    
    # Barkod kolonu için EAN/UPC uyuşması
    alloc_barkod_col = "EAN/UPC" if "EAN/UPC" in df_orders.columns else ("Barkod" if "Barkod" in df_orders.columns else None)
    if alloc_barkod_col is None:
        # Eğer ikisi de yoksa, içinde EAN veya UPC geçen ilk kolonu seç
        for col in df_orders.columns:
            if "ean" in col.lower() or "upc" in col.lower() or "barkod" in col.lower():
                alloc_barkod_col = col
                break
                
    # Tahsis Edilen Adet kolonu için gelişmiş arama listesi
    alokasyon_adet_col = None
    for col in df_orders.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ["tahsis", "alokasyon", "allocated", "tahsis edilen adet", "tahsis miktarı"]):
            alokasyon_adet_col = col
            break
            
    # Kalan Adet kolonu için gelişmiş arama listesi
    kalan_adet_col = None
    for col in df_orders.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ["kalan adet", "kalan", "remaining", "kalan miktar", "bakiye"]):
            kalan_adet_col = col
            break
            
    # --- GÜVENLİK AĞI (FALLBACK): Eğer hala bulunamadıysa varsayılan ata ---
    if alokasyon_adet_col is None:
        alokasyon_adet_col = "Tahsis Edilen Adet" if "Tahsis Edilen Adet" in df_orders.columns else "Tahsis Miktarı"
    if kalan_adet_col is None:
        kalan_adet_col = "Kalan adet" if "Kalan adet" in df_orders.columns else "Kalan Adet"
    if alloc_barkod_col is None:
        alloc_barkod_col = "EAN/UPC"

    # Kolon Kontrolleri
    orders_ok = siparis_material_col in df_orders.columns and "Sipariş Miktarı" in df_orders.columns and alokasyon_adet_col in df_orders.columns and kalan_adet_col in df_orders.columns and alloc_barkod_col in df_orders.columns
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
        df_orders[siparis_material_col] = malzeme_kodunu_temizle(df_orders[siparis_material_col])
        df_orders[alloc_barkod_col] = gelismis_barkod_temizle(df_orders[alloc_barkod_col])
        
        df_catalog[katalog_material_col] = malzeme_kodunu_temizle(df_catalog[katalog_material_col])
        df_catalog[katalog_ean_col] = gelismis_barkod_temizle(df_catalog[katalog_ean_col])
        
        df_stock[stok_material_col] = malzeme_kodunu_temizle(df_stock[stok_material_col])
        df_stock[stok_net_avail_col] = pd.to_numeric(df_stock[stok_net_avail_col], errors='coerce').fillna(0)

        # --- BARKOD BAZLI STOK HESAPLAMA ---
        df_stock_grouped = df_stock.groupby(stok_material_col)[stok_net_avail_col].sum().reset_index()
        df_cat_bridge = df_catalog[[katalog_material_col, katalog_ean_col]].dropna().drop_duplicates()
        df_merged_stock = pd.merge(df_cat_bridge, df_stock_grouped, on=katalog_material_col, how="inner")
        
        df_barcode_stock_sum = df_merged_stock.groupby(katalog_ean_col)[stok_net_avail_col].sum().reset_index()
        df_barcode_stock_sum.rename(columns={katalog_ean_col: "EAN_Köprü"}, inplace=True)

        # --- SİPARİŞ VE STOK BİRLEŞTİRME ---
        df_final = pd.merge(df_orders, df_barcode_stock_sum, left_on=alloc_barkod_col, right_on="EAN_Köprü", how="left")
        df_final[stok_net_avail_col] = df_final[stok_net_avail_col].fillna(0)
        
        # --- BARKOD BAZLI TOPLAM KALAN ALOKASYON VE BOŞTAKİ STOK HESABI ---
        df_final["Toplam Kalan Alokasyon"] = df_final.groupby(alloc_barkod_col)[kalan_adet_col].transform("sum")
        
        df_final["Boştaki Stok"] = df_final[stok_net_avail_col] - df_final["Toplam Kalan Alokasyon"]
        df_final["Boştaki Stok"] = df_final["Boştaki Stok"].apply(lambda x: max(0, x))

        # --- AUDIT REVİZYON ALGORİTMASI ---
        def audit_karar_motoru(row):
            kalan = row[kalan_adet_col]
            boştaki_stok = row["Boştaki Stok"]
            
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
        
        display_cols = []
        possible_cols = {
            'Müşteri Adı': 'Müşteri Adı',
            siparis_material_col: siparis_material_col,
            alloc_barkod_col: alloc_barkod_col,
            'Sipariş Miktarı': 'Sipariş Miktarı',
            alokasyon_adet_col: alokasyon_adet_col,
            kalan_adet_col: kalan_adet_col,
            stok_net_avail_col: stok_net_avail_col,
            'Boştaki Stok': 'Boştaki Stok',
            'Denetim Durumu': 'Denetim Durumu',
            'Önerilen Aksiyon Planı': 'Önerilen Aksiyon Planı'
        }
        for col_key, col_val in possible_cols.items():
            if col_val in df_filtered.columns:
                display_cols.append(col_val)
        
        def color_audit_durumu(val):
            if "Boşta Stok" in val:
                return 'color: #1E90FF; font-weight: bold;'
            else:
                return 'color: green; font-weight: bold;'

        st.dataframe(
            df_filtered[display_cols].style.format({
                'Sipariş Miktarı': '{:,.0f}' if 'Sipariş Miktarı' in df_filtered.columns else '{}',
                alokasyon_adet_col: '{:,.0f}' if alokasyon_adet_col in df_filtered.columns else '{}',
                kalan_adet_col: '{:,.0f}' if kalan_adet_col in df_filtered.columns else '{}',
                stok_net_avail_col: '{:,.0f}' if stok_net_avail_col in df_filtered.columns else '{}',
                'Boştaki Stok': '{:,.0f}' if 'Boştaki Stok' in df_filtered.columns else '{}'
            }).map(color_audit_durumu, subset=['Denetim Durumu'] if 'Denetim Durumu' in df_filtered.columns else []),
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
            st.error(f"❌ Sipariş dosyasında '{siparis_material_col}', 'Sipariş Miktarı', '{alokasyon_adet_col}', '{kalan_adet_col}' ve '{alloc_barkod_col}' olmalı. (Yüklenen sayfa: Sheet1)")
        if not catalog_ok:
            st.error(f"❌ Katalog dosyasında '{katalog_material_col}' ve '{katalog_ean_col}' olmalı.")
        if not stock_ok:
            st.error(f"❌ Stok dosyasında '{stok_material_col}' ve '{stok_net_avail_col}' olmalı.")
else:
    st.info("💡 Lütfen sol menüden 'Siparişler', 'Katalog' ve 'Stok' excel dosyalarını yükleyin. Denetim analizleri arka planda anlık olarak tamamlanacaktır.")
