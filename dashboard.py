import streamlit as st
import pandas as pd
import io
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape
import numpy as np
import re

# ------------------------------
# Funzione: genera PDF tabellare
# ------------------------------


def generate_pdf(df, title="Report"):
    buffer = io.BytesIO()

    # Landscape se troppe colonne
    if len(df.columns) > 5:
        pagesize = landscape(A4)
    else:
        pagesize = A4

    doc = SimpleDocTemplate(buffer, pagesize=pagesize, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []

    styles = getSampleStyleSheet()
    title_style = styles['Title']
    normal_style = styles['Normal']
    normal_style.wordWrap = 'LTR'   # wrapping classico (non CJK)
    header_style = styles['Heading5']

    # Titolo
    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 12))

    # Header
    data = [[Paragraph(f"<nobr>{str(col).replace("_", " ")}</nobr>", styles['Normal']) for col in df.columns]]

    # Dati
    for _, row in df.iterrows():
        formatted_row = []
        for col, cell in zip(df.columns, row):
            if isinstance(cell, (int, float, np.number)):
                formatted_row.append(Paragraph(f"{cell:.2f}", normal_style))
            else:
                formatted_row.append(Paragraph(str(cell), normal_style))
        data.append(formatted_row)

    # Calcolo larghezze colonne
    page_width = pagesize[0] - 40
    col_widths = []
    for col in df.columns:
        texts = [str(col)] + df[col].astype(str).tolist()
        max_width = max(stringWidth(t, "Helvetica", 8) for t in texts) + 15
        col_widths.append(max_width)

    total_width = sum(col_widths)
    if total_width > page_width:
        scale = page_width / total_width
        col_widths = [w * scale for w in col_widths]

    table = Table(data, colWidths=col_widths, repeatRows=1)

    # Font size dinamico
    if len(df.columns) <= 6:
        font_size = 8
    elif len(df.columns) <= 10:
        font_size = 7
    else:
        font_size = 6

    # Stile tabella
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#d3d3d3")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), font_size),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('WORDWRAP', (0, 0), (-1, 0), 'CJK')
    ])
    table.setStyle(style)

    elements.append(table)
    doc.build(elements)

    buffer.seek(0)
    return buffer


if "editor_file" not in st.session_state:
    st.session_state.editor_file = None
if "consolidato_file" not in st.session_state:
    st.session_state.consolidato_file = None
if "editor_df" not in st.session_state:
    st.session_state.editor_df = None
if "consolidato_df" not in st.session_state:
    st.session_state.consolidato_df = None

# ------------------------------
# Sidebar: scelta pagina
# ------------------------------
st.sidebar.title("📂 Navigazione")
page = st.sidebar.radio("Vai a:", ["📊 Capability", "📈 Consolidato"])

if page == "📊 Capability":
    st.title("Editor Excel")

    uploaded_file = st.file_uploader("📂 Carica un file Excel Capability per modificarlo", 
                                     type=["xlsx", "xls"], 
                                     key="editor")
    
    if uploaded_file is not None:
        st.session_state.editor_file = uploaded_file

    if st.session_state.editor_file is not None:
        xls = pd.ExcelFile(st.session_state.editor_file)
        sheet_names = xls.sheet_names

        selected_sheet = st.selectbox("Seleziona il foglio da visualizzare", sheet_names)
        df = pd.read_excel(st.session_state.editor_file, sheet_name=selected_sheet)

        st.session_state.editor_df = df.copy()

        date_cols = ["Data inizio collaborazione\n(gg/mm/aaaa)", "Data fine collaborazione\n(gg/mm/aaaa)"]
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        for col in df.select_dtypes(include=["datetime64[ns]"]).columns:
            df[col] = df[col].dt.strftime("%d/%m/%Y")

        if "Commenti" in df.columns:
            df["Commenti"] = df["Commenti"].astype("str")

        st.write(f"Modifica il foglio **{selected_sheet}**:")
        edited_df = st.data_editor(
            df.style.format({col: "{:.2f}" for col in df.select_dtypes(include=["number"]).columns}), 
            num_rows="dynamic", 
            use_container_width=True
            )

        # Campo per inserire il nome del file Excel
        default_editor_excel = "file_modificato.xlsx"
        editor_excel_name = st.text_input("📊 Inserisci il nome del file Excel da scaricare:", value=default_editor_excel, key="editor_excel_name")

        if st.button("📥 Scarica file aggiornato (Excel)"):
            all_sheets = {sheet: pd.read_excel(st.session_state.editor_file, sheet_name=sheet) for sheet in sheet_names}
            all_sheets[selected_sheet] = edited_df

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet, data in all_sheets.items():
                    data.to_excel(writer, sheet_name=sheet, index=False)
            output.seek(0)

            st.download_button(
                label="⬇️ Scarica Excel modificato",
                data=output,
                file_name=editor_excel_name if editor_excel_name.endswith(".xlsx") else editor_excel_name + ".xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ------------------------------
# PAGINA ANALISI CONSOLIDATO
# ------------------------------
elif page == "📈 Consolidato":
    st.title("📈 Analisi consolidato mensile")

    with st.expander("📖 Guida: come preparare il file consolidato"):
        st.markdown("""
        **Passo 1:** Aprire il file `T05-Consolidato`, scegliere il foglio di interesse, selezionare tutto 
        (Ctrl+A o dall’angolo in alto a sinistra) e copiare (Ctrl+C).

        **Passo 2:** Aprire un nuovo file Excel vuoto ed incollare i dati (Ctrl+V).

        **Passo 3:** Rinominare il foglio con lo stesso nome del file originale e salvare il file con un nuovo nome 
        (ad esempio `Consolidato_MeseCorrente.xlsx`).

        **Passo 4:** Caricare qui il nuovo file salvato.  
        _(Questo evita problemi dovuti alle pivot presenti negli altri fogli del file principale.)_
        """)

    consolidato_file = st.file_uploader("📂 Inserire file consolidato mensile", 
                                        type=["xlsx", "xls"], 
                                        key="consolidato")
    
    if consolidato_file is not None:
        st.session_state.consolidato_file = consolidato_file

    if st.session_state.consolidato_file is not None:
        df_cons = pd.read_excel(st.session_state.consolidato_file, sheet_name=0)
        st.session_state.consolidato_df = df_cons.copy()

        required_cols = ["Supplier", "FTEs"]
        if not all(col in df_cons.columns for col in required_cols):
            st.error("Il file consolidato deve contenere almeno le colonne 'Supplier' e 'FTEs'.")
        else:
            def clean_fte(x):
                if pd.isna(x):
                    return np.nan
                if isinstance(x, (int, float, np.number)):
                    return float(x)
                s = str(x).strip().replace(",", ".")
                s = re.sub(r"[^0-9.-]", "", s)
                if s in ["", ".", "-"]:
                    return np.nan
                try:
                    return float(s)
                except:
                    return np.nan

            df_cons["Supplier"] = df_cons["Supplier"].astype(str).str.strip()
            df_cons["FTEs_clean"] = df_cons["FTEs"].apply(clean_fte)
            df_cons["Supplier_norm"] = df_cons["Supplier"].str.split().str.join(" ")

            colonna_ordinamento = "Giugno '25 - In/Out"
            valori_in = ["IN", "IN_dd", "IN_nb", "IN_rnm", "TBV (in)"]
            df_cons = df_cons[df_cons[colonna_ordinamento].isin(valori_in)].copy()

            # Aggregazione per supplier
            agg_cons = (
                df_cons.groupby("Supplier_norm", dropna=False)["FTEs_clean"]
                .sum(min_count=1)
                .reset_index()
                .rename(columns={"FTEs_clean": "FTEs_total"})
                .sort_values("Supplier_norm")
            )

            # Filtro supplier con FTE tra 0 e 3
            selected_suppliers_norm = agg_cons[
                (agg_cons["FTEs_total"] >= 0) & (agg_cons["FTEs_total"] <= 3)
            ]["Supplier_norm"].tolist()

            rows_selected = (
                df_cons[df_cons["Supplier_norm"].isin(selected_suppliers_norm)]
                .copy()
                .merge(agg_cons, on="Supplier_norm", how = "left")
                .sort_values(["Supplier_norm", "L1: Capability/Function"])
            )

            # Mostra risultati aggregati
            st.subheader("📊 Totali FTE per Supplier")
            st.dataframe(agg_cons.style.format({"FTEs_total": "{:.2f}"}))

            if not rows_selected.empty:
                st.subheader("📋 Supplier con FTE totale tra 0 e 3 (con Capability/Function)")

                # Controlla se esiste la colonna 'L1: Capability/Function'
                capability_col = "L1: Capability/Function"
                res_id_col = "RES ID (SNow)"

                display_cols = ["Supplier"]

                if capability_col in rows_selected.columns:
                    display_cols.append(capability_col)
                if res_id_col in rows_selected.columns:
                    display_cols.append(res_id_col)

                display_cols.extend(["FTEs", "FTEs_total"])

                st.dataframe(
                    rows_selected[display_cols]
                    .style.format({"FTEs_clean": "{:.2f}", "FTEs_total": "{:.2f}"})
                )

                # Download Excel
                default_excel_name = ".xlsx"
                excel_name = st.text_input("📊 Inserisci il nome del file Excel completo:", value=default_excel_name)

                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    df_cons.to_excel(writer, sheet_name="Originale", index=False)
                    agg_cons.to_excel(writer, sheet_name="Supplier_FTE_totals", index=False)
                    rows_selected.to_excel(writer, sheet_name="FTE_0_3_rows", index=False)
                excel_buffer.seek(0)

                st.download_button(
                    label="⬇️ Scarica risultato in Excel",
                    data=excel_buffer,
                    file_name=excel_name if excel_name.endswith(".xlsx") else excel_name + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


                # Download Excel (solo tabella Supplier con FTE 0-3)
                default_excel_single = ".xlsx"
                excel_name_single = st.text_input("📊 Inserisci il nome del file Excel (solo tabella FTE 0-3):", value=default_excel_single)

                excel_buffer_single = io.BytesIO()
                with pd.ExcelWriter(excel_buffer_single, engine="openpyxl") as writer:
                    rows_selected[display_cols].to_excel(writer, sheet_name="FTE_0_3_rows", index=False)
                excel_buffer_single.seek(0)

                st.download_button(
                    label="⬇️ Scarica SOLO tabella Supplier FTE 0-3 in Excel",
                    data=excel_buffer_single,
                    file_name=excel_name_single if excel_name_single.endswith(".xlsx") else excel_name_single + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # Download PDF
                default_name = ".pdf"
                pdf_name = st.text_input("📄 Inserisci il nome del file PDF da scaricare: (Supplier_FTE_0-3_MeseCorrente)", value=default_name)

                pdf_buffer = generate_pdf(
                    rows_selected[display_cols],
                    title="Supplier con FTE tra 0 e 3 (Consolidato)"
                )
                st.download_button(
                    label="⬇️ Scarica risultato in PDF",
                    data=pdf_buffer,
                    file_name=pdf_name if pdf_name.endswith(".pdf") else pdf_name + ".pdf",
                    mime="application/pdf",
                )
            else:
                st.info("Nessun supplier con FTE totale tra 0 e 3 trovato nel consolidato.")