"""
iddia_tahmini.py — Balıkesir Üniversitesi 2026 Program Sırası Tahmin Uygulaması
================================================================================
Çalıştırma: streamlit run iddia_tahmini.py
"""

from __future__ import annotations

import math
import warnings
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# AYARLAR
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).resolve().parent
EXCEL_PATH = next(
    (p for p in [
        BASE_DIR / "son_3_yil_program_basari_istatistigi.xlsx",
        BASE_DIR / "data" / "son_3_yil_program_basari_istatistigi.xlsx",
    ] if p.exists()),
    None,
)
YEAR_COLS  = ["2023", "2024", "2025"]
PRED_YEAR  = 2026
YEAR_NUMS  = [2023, 2024, 2025]

# Renkler
C_PRIMARY   = "#1F3A6E"
C_ACCENT    = "#008478"
C_SAFE      = "#065F46"
C_SAFE_BG   = "#D1FAE5"
C_WARN      = "#92400E"
C_WARN_BG   = "#FEF3C7"
C_DANGER    = "#991B1B"
C_DANGER_BG = "#FEE2E2"
C_NEUTRAL   = "#4A5568"
C_PRED      = "#E07B00"          # turuncu — tahmin noktası
C_BG        = "#FFFFFF"
C_CARD      = "#F8FAFC"
C_BORDER    = "#E5E7EB"


# ─────────────────────────────────────────────────────────────────────────────
# VERİ YÜKLEMESİ
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data() -> pd.DataFrame:
    if EXCEL_PATH is None:
        st.error("🚨 Excel dosyası bulunamadı! `son_3_yil_program_basari_istatistigi.xlsx` repo kökünde olmalı.")
        st.stop()
    df = pd.read_excel(EXCEL_PATH, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df["OKUL ADI"] = df["OKUL ADI"].ffill()
    df["PROGRAM ADI"] = df["PROGRAM ADI"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    for col in YEAR_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["PROGRAM ADI"])
    df = df[df["PROGRAM ADI"] != "nan"]
    return df.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAHMİN ALGORİTMASI
# ─────────────────────────────────────────────────────────────────────────────
def predict_2026(v23: Optional[float], v24: Optional[float], v25: Optional[float]):
    """
    Lineer regresyon ile 2026 tahminini hesaplar.
    Dönüş: (pred_value, r2_score, veri_sayisi, trend_per_year)
    """
    pts = [(y, v) for y, v in zip(YEAR_NUMS, [v23, v24, v25])
           if v is not None and not (isinstance(v, float) and math.isnan(v))]

    if len(pts) == 0:
        return None, None, 0, None

    xs = np.array([p[0] for p in pts], dtype=float)
    ys = np.array([p[1] for p in pts], dtype=float)

    if len(pts) == 1:
        return int(ys[0]), None, 1, None

    # OLS
    xm, ym = xs.mean(), ys.mean()
    denom = np.dot(xs - xm, xs - xm)
    slope = np.dot(xs - xm, ys - ym) / denom if denom != 0 else 0.0
    intercept = ym - slope * xm
    pred = slope * PRED_YEAR + intercept

    # R²
    if len(pts) >= 3:
        y_hat = slope * xs + intercept
        ss_res = np.sum((ys - y_hat) ** 2)
        ss_tot = np.sum((ys - ym) ** 2)
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else 1.0
    else:
        r2 = None

    return max(1, int(round(pred))), r2, len(pts), float(slope)


def confidence_label(r2, n_pts):
    """R² ve veri sayısına göre güven etiketi."""
    if n_pts == 1:
        return "Veri Yetersiz", "⚫"
    if n_pts == 2:
        return "Düşük Güven", "🟡"
    # 3 nokta
    if r2 is None:
        return "Orta Güven", "🟡"
    if r2 >= 0.95:
        return "Yüksek Güven", "🟢"
    if r2 >= 0.70:
        return "Orta Güven", "🟡"
    return "Düşük Güven", "🔴"


def risk_of_student(student_rank: int, predicted_cutoff: int):
    """
    student_rank  : öğrencinin tahmini YKS sıralaması (düşük = iyi)
    predicted_cutoff: programın 2026 son kabul sırası tahmini
    Dönüş: (risk_label, margin_pct, badge, row_bg)
    """
    if predicted_cutoff is None:
        return "Tahmin Yok", None, "⚫", "#F3F4F6"

    margin_pct = (predicted_cutoff - student_rank) / predicted_cutoff * 100

    if margin_pct >= 15:
        return "Güvenli", margin_pct, "🟢", C_SAFE_BG
    if margin_pct >= 0:
        return "Dikkatli", margin_pct, "🟡", C_WARN_BG
    if margin_pct >= -10:
        return "Riskli", margin_pct, "🔴", "#FEE2E2"
    return "Ulaşılamaz", margin_pct, "⛔", "#F9FAFB"


# ─────────────────────────────────────────────────────────────────────────────
# GRAFİK
# ─────────────────────────────────────────────────────────────────────────────
def build_trend_chart(
    program: str,
    v23, v24, v25,
    pred_2026: Optional[int],
    student_rank: Optional[int],
) -> go.Figure:
    """Geçmiş + tahmin + öğrenci çizgisini gösteren Plotly grafiği."""

    hist_x, hist_y = [], []
    for y, v in zip(YEAR_COLS, [v23, v24, v25]):
        if v is not None and not (isinstance(v, float) and math.isnan(v)):
            hist_x.append(y)
            hist_y.append(int(v))

    fig = go.Figure()

    # Geçmiş çizgi
    fig.add_trace(go.Scatter(
        x=hist_x, y=hist_y,
        mode="lines+markers",
        name="Gerçek Sıra",
        line=dict(color=C_PRIMARY, width=3),
        marker=dict(size=9, color=C_PRIMARY),
        hovertemplate="%{x}: <b>%{y:,.0f}</b><extra></extra>",
    ))

    # Tahmin noktasına köprü çizgi (kesik)
    if hist_x and pred_2026 is not None:
        fig.add_trace(go.Scatter(
            x=[hist_x[-1], str(PRED_YEAR)],
            y=[hist_y[-1], pred_2026],
            mode="lines",
            name="Tahmin Trendi",
            line=dict(color=C_PRED, width=2, dash="dash"),
            showlegend=False,
            hoverinfo="skip",
        ))

    # Tahmin noktası
    if pred_2026 is not None:
        fig.add_trace(go.Scatter(
            x=[str(PRED_YEAR)], y=[pred_2026],
            mode="markers",
            name=f"2026 Tahmini",
            marker=dict(size=13, color=C_PRED, symbol="diamond",
                        line=dict(color="white", width=2)),
            hovertemplate=f"2026 Tahmini: <b>{pred_2026:,}</b><extra></extra>".replace(",", "."),
        ))

    # Öğrenci sıralaması yatay çizgisi
    if student_rank is not None:
        all_y = hist_y + ([pred_2026] if pred_2026 else [])
        x_vals = hist_x + ([str(PRED_YEAR)] if pred_2026 else [])
        fig.add_hline(
            y=student_rank,
            line=dict(color="#DC2626", width=2, dash="dot"),
            annotation_text=f"  Senin sıran: {student_rank:,}".replace(",", "."),
            annotation_font=dict(color="#DC2626", size=12),
            annotation_position="top left",
        )

    # Y ekseni ters (düşük sıra = iyi = yukarıda)
    all_vals = [v for v in hist_y + ([pred_2026] if pred_2026 else []) + ([student_rank] if student_rank else [])
                if v is not None]
    if all_vals:
        y_min = min(all_vals)
        y_max = max(all_vals)
        pad = (y_max - y_min) * 0.18 or y_max * 0.1
        fig.update_yaxes(range=[y_max + pad, max(1, y_min - pad)], autorange=False)

    fig.update_layout(
        title=dict(
            text=f"<b>{program}</b><br><span style='font-size:12px;color:{C_NEUTRAL};'>Başarı Sırası Trendi</span>",
            font=dict(size=16, color=C_PRIMARY),
            x=0,
        ),
        paper_bgcolor=C_BG,
        plot_bgcolor=C_CARD,
        font=dict(family="Inter, sans-serif", color=C_NEUTRAL),
        xaxis=dict(
            gridcolor=C_BORDER, linecolor=C_BORDER,
            title="Yıl",
        ),
        yaxis=dict(
            gridcolor=C_BORDER, linecolor=C_BORDER,
            title="Başarı Sırası (düşük = iyi)",
            tickformat=",",
        ),
        legend=dict(orientation="h", y=-0.2),
        margin=dict(t=80, b=60, l=10, r=10),
        height=380,
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# YARDIMCI: sayı formatı
# ─────────────────────────────────────────────────────────────────────────────
def fmt(v) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    return f"{int(v):,}".replace(",", ".")


def pct_badge(pct: Optional[float]) -> str:
    if pct is None:
        return "—"
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:.1f}%"


# ─────────────────────────────────────────────────────────────────────────────
# SAYFA YAPISI
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="İddia Tahmini — BAUN 2026",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Genel stil
st.markdown("""
<style>
    .stApp { font-family: Inter, sans-serif; }
    div[data-testid="metric-container"] { background:#F8FAFC; border-radius:8px; padding:12px; }
    .risk-safe    { background:#D1FAE5; color:#065F46; border-radius:6px; padding:2px 8px; font-weight:700; }
    .risk-dikkat  { background:#FEF3C7; color:#92400E; border-radius:6px; padding:2px 8px; font-weight:700; }
    .risk-riskli  { background:#FEE2E2; color:#991B1B; border-radius:6px; padding:2px 8px; font-weight:700; }
    .risk-nope    { background:#F3F4F6; color:#6B7280; border-radius:6px; padding:2px 8px; font-weight:700; }
</style>
""", unsafe_allow_html=True)


# ── Başlık ──────────────────────────────────────────────────────────────────
st.markdown(
    f'<h1 style="color:{C_PRIMARY}; margin-bottom:0;">🎯 2026 İddia Tahmini</h1>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<p style="color:{C_NEUTRAL}; margin-top:4px; font-size:15px;">'
    'Tahmini YKS sıran ile 2026 yılında hangi BAUN programlarına girebileceğini öğren.</p>',
    unsafe_allow_html=True,
)
st.markdown("---")

# ── Veri yükle ───────────────────────────────────────────────────────────────
df = load_data()

# ── Tahminleri hesapla (cache'li) ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def compute_predictions(df_json: str) -> pd.DataFrame:
    df_in = pd.read_json(df_json, orient="records")
    rows = []
    for _, row in df_in.iterrows():
        v23 = row.get("2023")
        v24 = row.get("2024")
        v25 = row.get("2025")
        pred, r2, n_pts, slope = predict_2026(
            None if pd.isna(v23) else float(v23),
            None if pd.isna(v24) else float(v24),
            None if pd.isna(v25) else float(v25),
        )
        clabel, cicon = confidence_label(r2, n_pts)
        rows.append({
            "OKUL ADI":     row["OKUL ADI"],
            "PROGRAM ADI":  row["PROGRAM ADI"],
            "2023":         None if pd.isna(v23) else int(v23),
            "2024":         None if pd.isna(v24) else int(v24),
            "2025":         None if pd.isna(v25) else int(v25),
            "TAHMİN_2026":  pred,
            "R2":           r2,
            "N_PTS":        n_pts,
            "SLOPE":        slope,
            "GÜVEN":        clabel,
            "GÜVEN_İKON":   cicon,
        })
    return pd.DataFrame(rows)


preds_df = compute_predictions(df.to_json(orient="records"))

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f'<h3 style="color:{C_PRIMARY};">⚙️ Tahmin Ayarları</h3>', unsafe_allow_html=True)

    student_rank = st.number_input(
        "📌 Tahmini YKS Sıralaman",
        min_value=1,
        max_value=3_000_000,
        value=300_000,
        step=1_000,
        help="YKS'den beklediğin başarı sırasını gir. Düşük = daha iyi.",
    )

    st.markdown("---")
    st.markdown("### 🔍 Filtrele")

    all_schools = ["Tümü"] + sorted(preds_df["OKUL ADI"].unique().tolist())
    sel_school = st.selectbox("Okul / Fakülte", all_schools)

    risk_filter = st.multiselect(
        "Risk Kategorisi",
        options=["Güvenli", "Dikkatli", "Riskli", "Ulaşılamaz", "Tahmin Yok"],
        default=["Güvenli", "Dikkatli", "Riskli"],
        help="Görmek istediğin risk kategorilerini seç.",
    )

    show_chart = st.checkbox("Seçili programa trend grafiği göster", value=True)

    st.markdown("---")
    st.markdown("""
**📖 Risk Kategorileri**
- 🟢 **Güvenli**: Sıran, tahmini kesim noktasından ≥15% daha iyi
- 🟡 **Dikkatli**: Sıran, tahmini kesim noktasının %0–15% altında
- 🔴 **Riskli**: Sıran, tahmini kesim noktasını %10 aştı
- ⛔ **Ulaşılamaz**: Sıran, tahmini kesim noktasının çok üstünde
""")
    st.caption(f"📊 Toplam **{len(preds_df)}** program | **{preds_df['OKUL ADI'].nunique()}** okul")


# ── Risk hesabını tabloya ekle ────────────────────────────────────────────────
def enrich(df_p: pd.DataFrame, s_rank: int) -> pd.DataFrame:
    df_p = df_p.copy()
    df_p["RİSK"], df_p["MARJ_%"], df_p["ROZET"], df_p["SATIR_BG"] = zip(*[
        risk_of_student(s_rank, r["TAHMİN_2026"])
        for _, r in df_p.iterrows()
    ])
    return df_p


enriched = enrich(preds_df, student_rank)

# ── Okul filtresi ─────────────────────────────────────────────────────────────
if sel_school != "Tümü":
    enriched = enriched[enriched["OKUL ADI"] == sel_school]

# ── Risk filtresi ─────────────────────────────────────────────────────────────
if risk_filter:
    enriched = enriched[enriched["RİSK"].isin(risk_filter)]


# ── Özet metrikler ────────────────────────────────────────────────────────────
st.markdown(f"### 📈 Sıralaman: **{student_rank:,}** — Sonuçlar".replace(",", "."))

full_enriched = enrich(preds_df, student_rank)
n_guvenli   = (full_enriched["RİSK"] == "Güvenli").sum()
n_dikkatli  = (full_enriched["RİSK"] == "Dikkatli").sum()
n_riskli    = (full_enriched["RİSK"] == "Riskli").sum()
n_nope      = (full_enriched["RİSK"] == "Ulaşılamaz").sum()

c1, c2, c3, c4 = st.columns(4)
c1.metric("🟢 Güvenli",    n_guvenli,   help="Sıran, tahminden ≥15% iyi")
c2.metric("🟡 Dikkatli",   n_dikkatli,  help="Sıran, tahminden %0–15% iyi")
c3.metric("🔴 Riskli",     n_riskli,    help="Sıran tahmini %10'a kadar aştı")
c4.metric("⛔ Ulaşılamaz", n_nope,      help="Sıran tahmininin çok üstünde")

st.markdown("---")

# ── Tablo ─────────────────────────────────────────────────────────────────────
if enriched.empty:
    st.info("Seçili filtrelerle eşleşen program bulunamadı.")
else:
    st.markdown(f"**{len(enriched)} program** gösteriliyor.")

    # Sıralama: önce güvenli, sonra dikkatli...
    risk_order = {"Güvenli": 0, "Dikkatli": 1, "Riskli": 2, "Ulaşılamaz": 3, "Tahmin Yok": 4}
    enriched = enriched.copy()
    enriched["_sira"] = enriched["RİSK"].map(risk_order).fillna(5)
    enriched = enriched.sort_values(["_sira", "MARJ_%"], ascending=[True, False])

    # Gösterim sütunları
    display_cols = {
        "ROZET":       "Risk",
        "OKUL ADI":    "Okul / Fakülte",
        "PROGRAM ADI": "Program",
        "2023":        "2023 Sırası",
        "2024":        "2024 Sırası",
        "2025":        "2025 Sırası",
        "TAHMİN_2026": "2026 Tahmini",
        "MARJ_%":      "Marj %",
        "GÜVEN_İKON":  "Güven",
    }
    show_df = enriched[list(display_cols.keys())].copy()
    show_df.columns = list(display_cols.values())

    # Sayısal sütunları formatla
    for col in ["2023 Sırası", "2024 Sırası", "2025 Sırası", "2026 Tahmini"]:
        show_df[col] = show_df[col].apply(lambda x: fmt(x) if x else "—")

    show_df["Marj %"] = show_df["Marj %"].apply(
        lambda x: pct_badge(x) if x is not None else "—"
    )

    st.dataframe(
        show_df.reset_index(drop=True),
        use_container_width=True,
        height=420,
        column_config={
            "Risk":         st.column_config.TextColumn("Risk", width="small"),
            "Okul / Fakülte": st.column_config.TextColumn("Okul / Fakülte", width="medium"),
            "Program":      st.column_config.TextColumn("Program", width="medium"),
            "2026 Tahmini": st.column_config.TextColumn("2026 Tahmini ⭐", width="medium"),
            "Marj %":       st.column_config.TextColumn("Marj %", width="small"),
            "Güven":        st.column_config.TextColumn("Güven", width="small"),
        },
    )


# ── Program detay grafiği ─────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 🔍 Program Detay Grafiği")

all_programs = sorted(preds_df["PROGRAM ADI"].unique().tolist())
if sel_school != "Tümü":
    prog_list = sorted(preds_df[preds_df["OKUL ADI"] == sel_school]["PROGRAM ADI"].unique().tolist())
else:
    prog_list = all_programs

if prog_list:
    sel_prog = st.selectbox("Program seç", prog_list, key="detail_prog")
    sel_school_for_prog = preds_df[preds_df["PROGRAM ADI"] == sel_prog]["OKUL ADI"].iloc[0]

    row = preds_df[preds_df["PROGRAM ADI"] == sel_prog].iloc[0]
    enr_row = enrich(preds_df[preds_df["PROGRAM ADI"] == sel_prog], student_rank).iloc[0]

    col_chart, col_info = st.columns([2.2, 1], gap="large")

    with col_chart:
        fig = build_trend_chart(
            program=sel_prog,
            v23=row["2023"], v24=row["2024"], v25=row["2025"],
            pred_2026=row["TAHMİN_2026"],
            student_rank=student_rank if show_chart else None,
        )
        st.plotly_chart(fig, use_container_width=True,
                        config={"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]})

    with col_info:
        # Başlık
        st.markdown(
            f'<div style="font-size:20px;font-weight:700;color:{C_PRIMARY};">{sel_school_for_prog}</div>'
            f'<div style="font-size:14px;color:{C_NEUTRAL};margin-bottom:12px;">{sel_prog}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(f'<hr style="border:none;border-top:1px solid {C_BORDER};margin:8px 0;">', unsafe_allow_html=True)

        # Geçmiş değerler
        for label, val in [("2023 Başarı Sırası", row["2023"]),
                           ("2024 Başarı Sırası", row["2024"]),
                           ("2025 Başarı Sırası", row["2025"]),
                           ("2026 Tahmini ⭐",     row["TAHMİN_2026"])]:
            color = C_PRED if "Tahmini" in label else C_PRIMARY
            st.markdown(
                f"""<div style="background:{C_CARD};border-left:4px solid {color};
                               padding:10px 14px;border-radius:6px;margin-bottom:8px;">
                    <div style="font-size:12px;color:{C_NEUTRAL};font-weight:600;">{label}</div>
                    <div style="font-size:22px;font-weight:700;color:{C_PRIMARY};">{fmt(val)}</div>
                </div>""",
                unsafe_allow_html=True,
            )

        # Risk bilgisi
        rozet = enr_row["ROZET"]
        risk  = enr_row["RİSK"]
        marj  = enr_row["MARJ_%"]
        bg    = enr_row["SATIR_BG"]

        st.markdown(
            f"""<div style="background:{bg};border-radius:8px;padding:12px 16px;margin-top:4px;">
                <div style="font-size:12px;color:{C_NEUTRAL};font-weight:600;">SENİN DURUMUN</div>
                <div style="font-size:24px;font-weight:700;color:{C_PRIMARY};">{rozet} {risk}</div>
                <div style="font-size:14px;color:{C_NEUTRAL};">
                    Marj: <b>{pct_badge(marj)}</b><br>
                    Güven: <b>{row['GÜVEN_İKON']} {row['GÜVEN']}</b>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )


# ─────────────────────────────────────────────────────────────────────────────
# AÇIKLAMA KUTUSU
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📖 Bu uygulama nasıl çalışır?"):
    st.markdown(f"""
**Tahmin Yöntemi**
- 2023, 2024, 2025 verilerinden **lineer regresyon** ile 2026 tahmini yapılır.
- 3 yıl verisi olan programlarda **R² güven skoru** hesaplanır.
- Yalnızca 2025 verisi olan yeni programlar "Veri Yetersiz" olarak işaretlenir.

**Sıra Yorumu**
- ÖSYM'de **küçük sayı = yüksek başarı** (daha rekabetçi program).
- Grafiklerde Y ekseni **ters** çevrilmiştir; yukarı gitmek = iyileşme.
- Tahmin noktası (turuncu ◆) sadece eğilim devam ederse geçerlidir.

**Risk Kategorileri**
| Kategori | Açıklama |
|----------|----------|
| 🟢 Güvenli | Sıran, 2026 tahmininden ≥ %15 daha iyi |
| 🟡 Dikkatli | Sıran, 2026 tahmininden %0–%15 daha iyi |
| 🔴 Riskli | Sıran, 2026 tahminini ≤ %10 aştı |
| ⛔ Ulaşılamaz | Sıran, 2026 tahmininin çok üzerinde |

> ⚠️ Bu tahminler **istatistiksel eğilime** dayanır. Gerçek sonuçlar kontenjan, tercih dağılımı
> ve diğer faktörlere göre değişebilir. Yalnızca yönlendirici amaçla kullanın.
""")

st.caption("🎓 BAUN Program Başarı Sırası — 2026 İddia Tahmini • Streamlit + Plotly")
