import io
import re
import zipfile
from typing import List, Tuple, Optional

import pandas as pd
import streamlit as st
import plotly.express as px

MAX_POINTS = 70
APP_TITLE = "Személyes és társas kompetenciák mérése"
FOOTER_TEXT = "Készítette: Sulyok István Ált. Iskola és AMI 2026. Pálfi Gergő"
CATEGORY_COLS = ["Művészeti képzés", "Sportoló", "Tanodás"]
CHART_COLORS = ["#1f77b4", "#ff7f0e"]  # kék, narancssárga
CHART_FONT = "DejaVu Sans"


def wrap_label(text, max_chars=18):
    """Hosszú tengelyfeliratok tördelése Plotly diagramokhoz."""
    text = str(text)
    words = text.split()
    lines = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "<br>".join(lines)


def safe_filename(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű_\-\.]", "", s)
    return s or "nev_nelkul"


def is_marked(value) -> bool:
    """A csoportoszlopokban lévő x / X / igen / 1 jelöléseket igaznak veszi."""
    if pd.isna(value):
        return False
    text = str(value).strip().lower()
    return text in {"x", "1", "i", "igen", "true", "t", "✓", "✔"}


def add_category_columns(df: pd.DataFrame, category_source: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Biztosítja, hogy minden időszak adatában meglegyenek a K-L-M csoportoszlopok."""
    out = df.copy()
    if category_source is not None:
        keys = ["Név", "Osztály"]
        available = [c for c in CATEGORY_COLS if c in category_source.columns]
        if available:
            src = category_source[keys + available].drop_duplicates(keys)
            out = out.merge(src, on=keys, how="left", suffixes=("", "_src"))
            for c in available:
                src_c = f"{c}_src"
                if src_c in out.columns:
                    out[c] = out[c].combine_first(out[src_c]) if c in out.columns else out[src_c]
                    out.drop(columns=[src_c], inplace=True)
    for c in CATEGORY_COLS:
        if c not in out.columns:
            out[c] = pd.NA
        out[c] = out[c].apply(is_marked)
    return out


def read_two_periods(xlsx_bytes: bytes) -> Tuple[pd.DataFrame, Optional[pd.DataFrame], List[str], List[str]]:
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    sheets = xls.sheet_names
    if len(sheets) < 1:
        raise ValueError("Az Excel fájlban nincs munkalap.")

    period_labels = [sheets[0]]
    df1_raw = pd.read_excel(xls, sheet_name=sheets[0])

    df2_raw = None
    if len(sheets) >= 2:
        period_labels.append(sheets[1])
        df2_raw = pd.read_excel(xls, sheet_name=sheets[1])

    if df1_raw.shape[1] < 3:
        raise ValueError("Az első munkalapon legalább 3 oszlop kell: Név, Osztály, és mérési területek.")

    name_col = df1_raw.columns[0]
    class_col = df1_raw.columns[1]
    df1_raw = df1_raw.rename(columns={name_col: "Név", class_col: "Osztály"})
    areas = [c for c in df1_raw.columns[2:] if c not in CATEGORY_COLS]
    df1 = add_category_columns(df1_raw[["Név", "Osztály"] + areas + [c for c in CATEGORY_COLS if c in df1_raw.columns]])

    df2 = None
    if df2_raw is not None:
        if df2_raw.shape[1] < 3:
            raise ValueError("A második munkalapon legalább 3 oszlop kell: Név, Osztály, és mérési területek.")
        name_col2 = df2_raw.columns[0]
        class_col2 = df2_raw.columns[1]
        df2_raw = df2_raw.rename(columns={name_col2: "Név", class_col2: "Osztály"})
        areas2 = [c for c in df2_raw.columns[2:] if c not in CATEGORY_COLS]

        common_areas = [a for a in areas if a in areas2]
        if not common_areas:
            raise ValueError("A két munkalap terület-oszlopai nem egyeznek (nincs közös terület).")

        areas = common_areas
        df1 = df1[["Név", "Osztály"] + areas + CATEGORY_COLS].copy()
        df2_base = df2_raw[["Név", "Osztály"] + areas + [c for c in CATEGORY_COLS if c in df2_raw.columns]].copy()
        df2 = add_category_columns(df2_base, category_source=df1[["Név", "Osztály"] + CATEGORY_COLS])

    return df1, df2, areas, period_labels


def to_long_percent(df: pd.DataFrame, period: str, areas: List[str]) -> pd.DataFrame:
    out = df.copy()
    for a in areas:
        out[a] = pd.to_numeric(out[a], errors="coerce")
    long = out.melt(
        id_vars=["Név", "Osztály"] + CATEGORY_COLS,
        value_vars=areas,
        var_name="Terület",
        value_name="Pont",
    )
    long["Időszak"] = period
    long["Százalék"] = (long["Pont"] / MAX_POINTS) * 100
    return long


def avg_view(long_df: pd.DataFrame, level: str) -> pd.DataFrame:
    if level == "Összes":
        return long_df.groupby(["Időszak", "Terület"], as_index=False)["Százalék"].mean()
    if level == "Csoport":
        # Több kiválasztott csoport esetén az aktuálisan szűrt tanulók közös átlaga.
        return long_df.groupby(["Időszak", "Terület"], as_index=False)["Százalék"].mean()
    return long_df.groupby(["Időszak", "Osztály", "Terület"], as_index=False)["Százalék"].mean()


def embed_footer_on_figure(fig, footer_text: str) -> None:
    fig.add_annotation(text=footer_text, x=0.5, y=-0.33, xref="paper", yref="paper", showarrow=False,
                       xanchor="center", yanchor="top", font=dict(size=12))
    fig.update_layout(margin=dict(b=180))


def make_bar_figure(data, area_order, period_order, title: str, embed_footer: bool = False, footer_text: str = FOOTER_TEXT):
    color_map = {period: CHART_COLORS[i % len(CHART_COLORS)] for i, period in enumerate(period_order)}
    fig = px.bar(
        data,
        x="Terület",
        y="Százalék",
        color="Időszak",
        color_discrete_map=color_map,
        category_orders={"Terület": area_order, "Időszak": period_order},
        barmode="group",
        range_y=[0, 110],
        labels={"Százalék": "Százalék (%)"},
        title=title,
    )
    fig.update_traces(texttemplate="%{y:.1f}%", textposition="outside", cliponaxis=False)
    fig.update_xaxes(
        tickmode="array",
        tickvals=area_order,
        ticktext=[wrap_label(a, 16) for a in area_order],
        tickfont=dict(size=9, family=CHART_FONT),
        automargin=True,
    )
    fig.update_yaxes(tickfont=dict(size=12, family=CHART_FONT))
    fig.update_layout(
        font=dict(family=CHART_FONT),
        legend_title_text="Időszak",
        xaxis_title="Mérési terület",
        title_font_size=18,
        margin=dict(t=55, b=150),
        uniformtext_minsize=8,
        uniformtext_mode="show",
        height=390,
    )
    if embed_footer:
        embed_footer_on_figure(fig, footer_text)
    return fig


def make_radar_figure(data, area_order, period_order, title: str, embed_footer: bool = False, footer_text: str = FOOTER_TEXT):
    color_map = {period: CHART_COLORS[i % len(CHART_COLORS)] for i, period in enumerate(period_order)}
    fig = px.line_polar(
        data,
        r="Százalék",
        theta="Terület",
        color="Időszak",
        color_discrete_map=color_map,
        category_orders={"Terület": area_order, "Időszak": period_order},
        line_close=True,
        range_r=[0, 100],
        labels={"Százalék": "Százalék (%)"},
        title=title,
    )
    fig.update_layout(font=dict(family=CHART_FONT), legend_title_text="Időszak", title_font_size=18, margin=dict(t=55, b=100), height=390)
    if embed_footer:
        embed_footer_on_figure(fig, footer_text)
    return fig


def figures_to_zip_png(fig_items, width=1400, height=950, scale=2) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, fig in fig_items:
            png_bytes = fig.to_image(format="png", width=width, height=height, scale=scale)
            zf.writestr(fname, png_bytes)
    return buf.getvalue()


def make_person_pdf(export_df: pd.DataFrame, area_order, period_order, title_prefix="Személyes és társas kompetencia", footer_text: str = FOOTER_TEXT, progress_callback=None) -> bytes:
    """Fekvő A4 PDF, keskeny margóval, egy oldalon két tanulóval.

    Fontos: ez a PDF-export már nem használ Kaleido/Chrome böngészőt.
    A diagramot közvetlenül ReportLab rajzolja, ezért nem jön elő a
    "Couldn't close or kill browser subprocess" hiba.
    """
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib import colors

    pdf_font = "DejaVuSans"
    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for fp in font_paths:
        try:
            pdfmetrics.registerFont(TTFont(pdf_font, fp))
            break
        except Exception:
            continue
    else:
        pdf_font = "Helvetica"

    page_w, page_h = landscape(A4)
    margin = 16
    gap = 12
    slot_w = (page_w - 2 * margin - gap) / 2
    slot_h = page_h - 2 * margin

    students = export_df[["Név", "Osztály"]].dropna().drop_duplicates().sort_values(["Osztály", "Név"]).to_records(index=False)
    total = len(students)
    if total == 0:
        raise ValueError("Nincs exportálható tanuló a szűrésben.")

    def wrap_plain(text, max_chars=15):
        words = str(text).split()
        lines, cur = [], ""
        for w in words:
            cand = w if not cur else cur + " " + w
            if len(cand) <= max_chars:
                cur = cand
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        return lines[:3]

    def draw_student_chart(c, one: pd.DataFrame, nev, osztaly, x0, y0, w, h):
        # keret és cím
        c.setStrokeColor(colors.lightgrey)
        c.rect(x0, y0, w, h, stroke=1, fill=0)
        c.setFont(pdf_font, 9)
        title = f"{title_prefix}: {nev} ({osztaly})"
        c.drawCentredString(x0 + w / 2, y0 + h - 14, title)

        # jelmagyarázat
        legend_y = y0 + h - 30
        for idx, period in enumerate(period_order):
            lx = x0 + 22 + idx * 78
            c.setFillColor(colors.HexColor(CHART_COLORS[idx % len(CHART_COLORS)]))
            c.rect(lx, legend_y - 6, 8, 8, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont(pdf_font, 6.5)
            c.drawString(lx + 12, legend_y - 5, str(period))

        plot_left = x0 + 30
        plot_right = x0 + w - 12

        # Kompaktabb PDF-diagram:
        # - nem csak a százalékos tengelyt rövidítjük,
        # - hanem a teljes oszlopdiagram rajzterületét kb. 25%-kal kisebbre vesszük,
        # - és közvetlenül a cím/jelmagyarázat alá húzzuk, hogy ne maradjon nagy üres sáv.
        original_plot_top = y0 + h - 48
        original_plot_bottom = y0 + 72
        original_plot_h = original_plot_top - original_plot_bottom
        plot_h = original_plot_h * 0.75
        plot_top = y0 + h - 56
        plot_bottom = plot_top - plot_h
        label_y = plot_bottom - 16

        plot_w = plot_right - plot_left

        # tengelyek és rács
        c.setStrokeColor(colors.lightgrey)
        c.setLineWidth(0.4)
        c.setFont(pdf_font, 5.5)
        for val in range(0, 101, 25):
            yy = plot_bottom + (val / 100) * plot_h
            c.line(plot_left, yy, plot_right, yy)
            c.setFillColor(colors.black)
            c.drawRightString(plot_left - 3, yy - 2, f"{val}%")
        c.setStrokeColor(colors.black)
        c.line(plot_left, plot_bottom, plot_right, plot_bottom)
        c.line(plot_left, plot_bottom, plot_left, plot_top)

        n_areas = max(len(area_order), 1)
        group_w = plot_w / n_areas
        inner_gap = group_w * 0.18
        n_periods = max(len(period_order), 1)
        bar_w = max((group_w - inner_gap) / n_periods * 0.72, 2)

        pivot = one.pivot_table(index="Terület", columns="Időszak", values="Százalék", aggfunc="mean").reindex(area_order).reindex(columns=period_order)

        for ai, area in enumerate(area_order):
            group_x = plot_left + ai * group_w
            center_x = group_x + group_w / 2
            bars_total_w = n_periods * bar_w
            start_x = center_x - bars_total_w / 2
            for pi, period in enumerate(period_order):
                val = pivot.loc[area, period] if area in pivot.index and period in pivot.columns else None
                if pd.isna(val):
                    continue
                val = max(0, min(float(val), 110))
                bh = min(val, 100) / 100 * plot_h
                bx = start_x + pi * bar_w
                c.setFillColor(colors.HexColor(CHART_COLORS[pi % len(CHART_COLORS)]))
                c.rect(bx, plot_bottom, bar_w * 0.86, bh, fill=1, stroke=0)
                c.setFillColor(colors.black)
                c.setFont(pdf_font, 4.8)
                c.drawCentredString(bx + bar_w * 0.43, plot_bottom + bh + 2, f"{val:.0f}%")

            # x felirat: kisebb, tördelve
            c.setFillColor(colors.black)
            c.setFont(pdf_font, 4.8)
            lines = wrap_plain(area, 13)
            for li, line in enumerate(lines):
                c.drawCentredString(center_x, label_y - li * 6, line)

        c.setFont(pdf_font, 6.5)
        c.setFillColor(colors.black)
        c.drawCentredString(x0 + w / 2, y0 + 6, footer_text)

    pdf_buf = io.BytesIO()
    c = canvas.Canvas(pdf_buf, pagesize=landscape(A4))

    for i, (nev, osztaly) in enumerate(students, start=1):
        if i > 1 and (i - 1) % 2 == 0:
            c.showPage()
        col = (i - 1) % 2
        x0 = margin + col * (slot_w + gap)
        y0 = margin
        one = export_df[(export_df["Név"] == nev) & (export_df["Osztály"] == osztaly)].copy()
        draw_student_chart(c, one, nev, osztaly, x0, y0, slot_w, slot_h)
        if progress_callback:
            progress_callback(i, total, f"PDF diagram rajzolása: {osztaly} – {nev}")

    c.save()
    pdf_buf.seek(0)
    return pdf_buf.getvalue()


def sport_status_label(value: bool) -> str:
    return "Sportoló" if bool(value) else "Nem sportoló"


def build_sport_compare_base(long_df: pd.DataFrame, selected_periods: List[str]) -> pd.DataFrame:
    """Sportoló / nem sportoló összehasonlító adatok előkészítése."""
    base = long_df[long_df["Időszak"].isin(selected_periods)].copy()
    base["Sport státusz"] = base["Sportoló"].apply(sport_status_label)
    return base


def sport_period_averages(base: pd.DataFrame) -> pd.DataFrame:
    return (
        base.groupby(["Időszak", "Sport státusz", "Terület"], as_index=False)
        .agg(
            Átlag=("Százalék", "mean"),
            Létszám=("Név", "nunique"),
        )
    )


def sport_overall_period_averages(base: pd.DataFrame) -> pd.DataFrame:
    """Területeken átlagolt összesített kompetenciaátlag."""
    return (
        base.groupby(["Időszak", "Sport státusz"], as_index=False)
        .agg(
            Átlag=("Százalék", "mean"),
            Létszám=("Név", "nunique"),
        )
    )


def sport_change_table(base: pd.DataFrame, start_period: str, end_period: str) -> pd.DataFrame:
    avg = sport_period_averages(base)
    start = avg[avg["Időszak"] == start_period][["Sport státusz", "Terület", "Átlag"]].rename(columns={"Átlag": f"{start_period} átlag"})
    end = avg[avg["Időszak"] == end_period][["Sport státusz", "Terület", "Átlag"]].rename(columns={"Átlag": f"{end_period} átlag"})
    out = start.merge(end, on=["Sport státusz", "Terület"], how="outer")
    out["Változás százalékpont"] = out[f"{end_period} átlag"] - out[f"{start_period} átlag"]
    return out


def sport_interaction_table(base: pd.DataFrame, start_period: str, end_period: str, area_order: List[str]) -> pd.DataFrame:
    changes = sport_change_table(base, start_period, end_period)
    piv = changes.pivot_table(index="Terület", columns="Sport státusz", values="Változás százalékpont", aggfunc="mean").reindex(area_order)
    for col in ["Sportoló", "Nem sportoló"]:
        if col not in piv.columns:
            piv[col] = pd.NA
    piv["Sportolói többletfejlődés"] = piv["Sportoló"] - piv["Nem sportoló"]
    piv = piv.reset_index()
    return piv[["Terület", "Sportoló", "Nem sportoló", "Sportolói többletfejlődés"]]


def make_sport_avg_figure(avg_df: pd.DataFrame, period: str, area_order: List[str], title: str):
    data = avg_df[avg_df["Időszak"] == period].copy()
    fig = px.bar(
        data,
        x="Terület",
        y="Átlag",
        color="Sport státusz",
        color_discrete_map={"Sportoló": CHART_COLORS[0], "Nem sportoló": CHART_COLORS[1]},
        category_orders={"Terület": area_order, "Sport státusz": ["Sportoló", "Nem sportoló"]},
        barmode="group",
        range_y=[0, 110],
        labels={"Átlag": "Átlag (%)"},
        title=title,
        hover_data={"Létszám": True, "Átlag": ":.1f"},
    )
    fig.update_traces(texttemplate="%{y:.1f}%", textposition="outside", cliponaxis=False)
    fig.update_xaxes(
        tickmode="array",
        tickvals=area_order,
        ticktext=[wrap_label(a, 16) for a in area_order],
        tickfont=dict(size=9, family=CHART_FONT),
        automargin=True,
    )
    fig.update_yaxes(tickfont=dict(size=12, family=CHART_FONT))
    fig.update_layout(
        font=dict(family=CHART_FONT),
        legend_title_text="Csoport",
        xaxis_title="Mérési terület",
        title_font_size=18,
        margin=dict(t=50, b=145),
        height=390,
    )
    return fig


def make_sport_change_figure(interaction_df: pd.DataFrame, area_order: List[str], title: str):
    data = interaction_df.melt(
        id_vars=["Terület"],
        value_vars=["Sportoló", "Nem sportoló", "Sportolói többletfejlődés"],
        var_name="Mutató",
        value_name="Százalékpont",
    )
    fig = px.bar(
        data,
        x="Terület",
        y="Százalékpont",
        color="Mutató",
        category_orders={"Terület": area_order, "Mutató": ["Sportoló", "Nem sportoló", "Sportolói többletfejlődés"]},
        barmode="group",
        labels={"Százalékpont": "Változás (százalékpont)"},
        title=title,
    )
    fig.update_traces(texttemplate="%{y:.1f}", textposition="outside", cliponaxis=False)
    fig.update_xaxes(
        tickmode="array",
        tickvals=area_order,
        ticktext=[wrap_label(a, 16) for a in area_order],
        tickfont=dict(size=9, family=CHART_FONT),
        automargin=True,
    )
    fig.update_layout(
        font=dict(family=CHART_FONT),
        legend_title_text="Mutató",
        xaxis_title="Mérési terület",
        title_font_size=18,
        margin=dict(t=50, b=145),
        height=390,
    )
    return fig


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Tömeges PNG/PDF exporthoz szükséges: `pip install kaleido reportlab`")

    uploaded = st.file_uploader("Excel fájl feltöltése (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.stop()

    try:
        df1, df2, areas, periods = read_two_periods(uploaded.getvalue())
    except Exception as e:
        st.error(f"Hiba a beolvasásnál: {e}")
        st.stop()

    long1 = to_long_percent(df1, periods[0], areas)
    long_all = long1.copy()
    if df2 is not None:
        long2 = to_long_percent(df2, periods[1], areas)
        long_all = pd.concat([long1, long2], ignore_index=True)

    area_order = areas[:]
    period_order = periods[:]

    left, _ = st.columns([1, 2])
    with left:
        view_mode = st.radio("Nézet", ["Egy tanuló", "Osztályátlag", "Csoportátlag", "Összes átlag"], index=0, key="view_mode")
        classes = sorted([c for c in long_all["Osztály"].dropna().unique().tolist()])
        selected_class = st.selectbox("Osztály szűrés", ["(mind)"] + classes, index=0, key="class_filter")
        selected_groups = st.multiselect("Csoport szűrés", CATEGORY_COLS, default=[], key="group_filter")
        group_logic = st.radio("Csoportszűrés módja", ["Bármelyik kiválasztott csoport", "Mindegyik kiválasztott csoport"], index=0, key="group_logic")
        show_table = st.checkbox("Táblázat mutatása (%)", value=True, key="show_table")

    filtered = long_all.copy()
    if selected_class != "(mind)":
        filtered = filtered[filtered["Osztály"] == selected_class]
    if selected_groups:
        mask = filtered[selected_groups].all(axis=1) if group_logic.startswith("Mindegyik") else filtered[selected_groups].any(axis=1)
        filtered = filtered[mask]

    if filtered.empty:
        st.warning("A megadott szűrőkkel nincs megjeleníthető adat.")
        st.stop()

    if view_mode == "Egy tanuló":
        names = sorted(filtered["Név"].dropna().unique().tolist())
        selected_name = st.selectbox("Tanuló", names, key="student_select")
        data = filtered[filtered["Név"] == selected_name].copy()
        chart_title = f"Személyes és társas kompetencia : {selected_name}"
    elif view_mode == "Osztályátlag":
        data = avg_view(filtered, "Osztály")
        if selected_class == "(mind)":
            cls = st.selectbox("Melyik osztály átlagát nézzük?", sorted(data["Osztály"].unique().tolist()), key="class_avg_pick")
            data = data[data["Osztály"] == cls].copy()
            chart_title = f"Személyes és társas kompetencia : {cls} átlag"
        else:
            chart_title = f"Személyes és társas kompetencia : {selected_class} átlag"
    elif view_mode == "Csoportátlag":
        data = avg_view(filtered, "Csoport")
        label = ", ".join(selected_groups) if selected_groups else "minden csoportjelölt / aktuális szűrés"
        chart_title = f"Személyes és társas kompetencia : {label} átlag"
    else:
        data = avg_view(filtered, "Összes")
        chart_title = "Személyes és társas kompetencia : Átlag"

    st.subheader("Területenkénti összehasonlítás (%)")
    st.plotly_chart(make_bar_figure(data, area_order, period_order, chart_title), width="stretch")
    st.caption(FOOTER_TEXT)

    st.subheader("Radar diagram (%)")
    st.plotly_chart(make_radar_figure(data, area_order, period_order, chart_title), width="stretch")
    st.caption(FOOTER_TEXT)

    if show_table:
        st.subheader("Százalékos táblázat (%)")
        pivot = data.pivot_table(index="Terület", columns="Időszak", values="Százalék", aggfunc="mean").reindex(area_order).reindex(columns=period_order)
        st.dataframe(pivot.style.format("{:.1f}%"), width="stretch")


    # ---------------- SPORTOLÓI ÖSSZEHASONLÍTÁS ----------------
    st.divider()
    st.subheader("Sportolói összehasonlítás")
    st.write(
        "Ez a rész a sportolók és nem sportolók átlagát, belső fejlődését és a "
        "sportolói többletfejlődést mutatja százalékpontban."
    )

    sport_source = long_all.copy()
    # A felső osztályszűrő itt is érvényes, mert így évfolyamra/osztályra külön is vizsgálható.
    if selected_class != "(mind)":
        sport_source = sport_source[sport_source["Osztály"] == selected_class]

    if sport_source.empty:
        st.info("A sportolói összehasonlításhoz nincs adat az aktuális osztályszűréssel.")
    else:
        sport_period_options = period_order[:]
        if len(sport_period_options) < 2:
            st.warning("A fejlődés és az interakció vizsgálatához legalább két mérési időszak szükséges.")
        else:
            scol1, scol2, scol3 = st.columns(3)
            with scol1:
                sport_start_period = st.selectbox(
                    "Kiinduló mérés",
                    sport_period_options,
                    index=0,
                    key="sport_start_period",
                )
            with scol2:
                default_end_idx = 1 if len(sport_period_options) > 1 else 0
                sport_end_period = st.selectbox(
                    "Záró mérés",
                    sport_period_options,
                    index=default_end_idx,
                    key="sport_end_period",
                )
            with scol3:
                sport_avg_period = st.selectbox(
                    "Átlagdiagram időszaka",
                    sport_period_options,
                    index=0,
                    key="sport_avg_period",
                )

            if sport_start_period == sport_end_period:
                st.warning("A kiinduló és záró mérés legyen különböző a változás kiszámításához.")
            else:
                selected_sport_periods = sorted(
                    {sport_start_period, sport_end_period, sport_avg_period},
                    key=lambda p: sport_period_options.index(p),
                )
                sport_base = build_sport_compare_base(sport_source, selected_sport_periods)
                sport_counts = (
                    sport_base[["Név", "Sport státusz"]]
                    .drop_duplicates()
                    .groupby("Sport státusz")["Név"]
                    .nunique()
                    .reindex(["Sportoló", "Nem sportoló"])
                    .fillna(0)
                    .astype(int)
                )

                m1, m2 = st.columns(2)
                m1.metric("Sportolók száma", int(sport_counts.get("Sportoló", 0)))
                m2.metric("Nem sportolók száma", int(sport_counts.get("Nem sportoló", 0)))

                if sport_counts.get("Sportoló", 0) == 0 or sport_counts.get("Nem sportoló", 0) == 0:
                    st.warning("Az összehasonlításhoz szükség van legalább egy sportolóra és egy nem sportolóra.")
                else:
                    sport_avg = sport_period_averages(sport_base)
                    st.plotly_chart(
                        make_sport_avg_figure(
                            sport_avg,
                            sport_avg_period,
                            area_order,
                            f"Sportolók és nem sportolók átlaga – {sport_avg_period}",
                        ),
                        width="stretch",
                    )

                    interaction = sport_interaction_table(sport_base, sport_start_period, sport_end_period, area_order)
                    st.plotly_chart(
                        make_sport_change_figure(
                            interaction,
                            area_order,
                            f"Fejlődés összehasonlítása: {sport_start_period} → {sport_end_period}",
                        ),
                        width="stretch",
                    )

                    overall = sport_overall_period_averages(sport_base)
                    overall_start = overall[overall["Időszak"] == sport_start_period][["Sport státusz", "Átlag"]].rename(columns={"Átlag": "Kiinduló átlag"})
                    overall_end = overall[overall["Időszak"] == sport_end_period][["Sport státusz", "Átlag"]].rename(columns={"Átlag": "Záró átlag"})
                    overall_change = overall_start.merge(overall_end, on="Sport státusz", how="outer")
                    overall_change["Változás százalékpont"] = overall_change["Záró átlag"] - overall_change["Kiinduló átlag"]

                    baseline_pivot = overall_change.set_index("Sport státusz")
                    sport_start = baseline_pivot.loc["Sportoló", "Kiinduló átlag"] if "Sportoló" in baseline_pivot.index else pd.NA
                    nonsport_start = baseline_pivot.loc["Nem sportoló", "Kiinduló átlag"] if "Nem sportoló" in baseline_pivot.index else pd.NA
                    sport_change = baseline_pivot.loc["Sportoló", "Változás százalékpont"] if "Sportoló" in baseline_pivot.index else pd.NA
                    nonsport_change = baseline_pivot.loc["Nem sportoló", "Változás százalékpont"] if "Nem sportoló" in baseline_pivot.index else pd.NA
                    baseline_diff = sport_start - nonsport_start
                    interaction_diff = sport_change - nonsport_change

                    q1, q2, q3 = st.columns(3)
                    q1.metric(
                        "Kiinduló különbség",
                        f"{baseline_diff:.1f} százalékpont",
                        help="Sportolók kiinduló átlaga mínusz nem sportolók kiinduló átlaga.",
                    )
                    q2.metric(
                        "Sportolók fejlődése",
                        f"{sport_change:.1f} százalékpont",
                        help="Sportolók záró átlaga mínusz kiinduló átlaga.",
                    )
                    q3.metric(
                        "Sportolói többletfejlődés",
                        f"{interaction_diff:.1f} százalékpont",
                        help="Sportolók fejlődése mínusz nem sportolók fejlődése.",
                    )

                    with st.expander("Részletes sportolói táblázatok"):
                        st.markdown("**Összesített átlag és változás**")
                        st.dataframe(overall_change.style.format({
                            "Kiinduló átlag": "{:.1f}%",
                            "Záró átlag": "{:.1f}%",
                            "Változás százalékpont": "{:.1f}",
                        }), width="stretch")

                        st.markdown("**Kompetenciaterületenkénti változás és többletfejlődés**")
                        st.dataframe(interaction.style.format({
                            "Sportoló": "{:.1f}",
                            "Nem sportoló": "{:.1f}",
                            "Sportolói többletfejlődés": "{:.1f}",
                        }), width="stretch")

                    st.info(
                        "Fejlesztési ötlet: az első kérdéshez érdemes lenne létszámot, szórást és "
                        "konfidenciaintervallumot is mutatni, mert így látszik, hogy az év eleji "
                        "sportoló–nem sportoló különbség mennyire stabil, nem csak az átlagokból adódó véletlen eltérés."
                    )


    st.divider()
    st.subheader("Tömeges letöltés")
    st.write("Az export az aktuális osztály- és csoportszűrést is figyelembe tudja venni.")

    export_scope = st.radio("Mit exportáljunk?", ["Minden tanuló (aktuális szűréssel)", "Minden tanuló (minden osztály és csoport)"], index=0, key="export_scope")
    export_df = long_all if export_scope.endswith("csoport)") else filtered
    only_bar_zip = st.checkbox("Csak oszlopdiagram letöltése a ZIP-ben", value=False, key="only_bar_zip")
    custom_footer_enabled = st.checkbox("Készítette szöveg átírása az exportban", value=False, key="custom_footer_enabled")
    export_footer_text = FOOTER_TEXT
    if custom_footer_enabled:
        export_footer_text = st.text_input("Készítette szöveg", value=FOOTER_TEXT, key="custom_footer_text").strip() or FOOTER_TEXT

    # A letöltőgombot nem szabad közvetlenül az elkészítő gomb belsejében hagyni,
    # mert a download_button kattintásakor a Streamlit újrafut, és eltűnhet a gomb.
    # Ezért az elkészült fájlokat session_state-ben tároljuk.
    st.session_state.setdefault("zip_ready", False)
    st.session_state.setdefault("zip_bytes", b"")
    st.session_state.setdefault("pdf_ready", False)
    st.session_state.setdefault("pdf_bytes", b"")
    st.session_state.setdefault("last_export_scope", export_scope)
    st.session_state.setdefault("last_only_bar_zip", only_bar_zip)
    st.session_state.setdefault("last_export_footer_text", export_footer_text)

    if (st.session_state["last_export_scope"] != export_scope
        or st.session_state["last_only_bar_zip"] != only_bar_zip
        or st.session_state["last_export_footer_text"] != export_footer_text):
        st.session_state["last_export_scope"] = export_scope
        st.session_state["last_only_bar_zip"] = only_bar_zip
        st.session_state["last_export_footer_text"] = export_footer_text
        st.session_state["zip_ready"] = False
        st.session_state["zip_bytes"] = b""
        st.session_state["pdf_ready"] = False
        st.session_state["pdf_bytes"] = b""

    zip_col, pdf_col = st.columns(2)
    with zip_col:
        if st.button("ZIP elkészítése PNG diagramokkal", type="primary", key="zip_make_btn"):
            st.session_state["zip_ready"] = False
            st.session_state["zip_bytes"] = b""
            try:
                fig_items = []
                export_names = export_df[["Név", "Osztály"]].dropna().drop_duplicates().sort_values(["Osztály", "Név"]).to_records(index=False)
                progress = st.progress(0)
                status = st.empty()
                total = len(export_names)
                for i, (nev, osztaly) in enumerate(export_names, start=1):
                    pct = int(i / max(total, 1) * 100)
                    status.write(f"Diagram előkészítése: {pct}% – {osztaly} – {nev} ({i}/{total})")
                    one = export_df[(export_df["Név"] == nev) & (export_df["Osztály"] == osztaly)].copy()
                    title = f"Személyes és társas kompetencia : {nev}"
                    base = f"{safe_filename(osztaly)}__{safe_filename(nev)}"
                    fig_items.append((f"{base}__oszlopdiagram.png", make_bar_figure(one, area_order, period_order, title, embed_footer=True, footer_text=export_footer_text)))
                    if not only_bar_zip:
                        fig_items.append((f"{base}__radar.png", make_radar_figure(one, area_order, period_order, title, embed_footer=True, footer_text=export_footer_text)))
                    progress.progress(min(i / max(total, 1), 1.0), text=f"{pct}%")
                status.write("ZIP csomag készítése…")
                st.session_state["zip_bytes"] = figures_to_zip_png(fig_items)
                st.session_state["zip_ready"] = True
                st.rerun()
            except Exception as e:
                st.error(f"Nem sikerült a ZIP export. Telepítve van a kaleido? Részletek: {e}")

        if st.session_state.get("zip_ready") and st.session_state.get("zip_bytes"):
            st.success("ZIP elkészült.")
            st.download_button(
                "Letöltés: diagramok ZIP",
                data=st.session_state["zip_bytes"],
                file_name="diagramok.zip",
                mime="application/zip",
                key="zip_download_btn",
            )

    with pdf_col:
        if st.button("PDF elkészítése nyomtatáshoz", type="primary", key="pdf_make_btn"):
            st.session_state["pdf_ready"] = False
            st.session_state["pdf_bytes"] = b""
            try:
                progress = st.progress(0, text="0%")
                status = st.empty()

                def pdf_progress(done, total, message):
                    pct = int(done / max(total, 1) * 100)
                    status.write(f"{message} – {pct}% ({done}/{total})")
                    progress.progress(min(done / max(total, 1), 1.0), text=f"{pct}%")

                st.session_state["pdf_bytes"] = make_person_pdf(export_df, area_order, period_order, footer_text=export_footer_text, progress_callback=pdf_progress)
                st.session_state["pdf_ready"] = True
                st.rerun()
            except Exception as e:
                st.error(f"Nem sikerült a PDF export. Telepítve van a reportlab? Részletek: {e}")

        if st.session_state.get("pdf_ready") and st.session_state.get("pdf_bytes"):
            st.success("PDF elkészült.")
            st.download_button(
                "Letöltés: nyomtatható PDF",
                data=st.session_state["pdf_bytes"],
                file_name="diagramok_nyomtathato.pdf",
                mime="application/pdf",
                key="pdf_download_btn",
            )


if __name__ == "__main__":
    main()
