import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

# -----------------------------
# 1. Carga y limpieza de datos
# -----------------------------
def load_and_clean(path: str) -> pd.DataFrame:
    # Detectar extensión
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in [".xls", ".xlsx"]:
        df = pd.read_excel(path)
    else:
        raise ValueError(f"Formato no soportado: {ext}")

    # Normalizar nombres de columnas (ajustar según archivo real)
    df = df.rename(columns={
        "FECHAVTO": "FECHAVTO",
        "NOFAC": "NOFAC",
        "PROVEEDOR": "PROVEEDOR",
        " SALDO ": "SALDO",   # si viene con espacios
        "SALDO": "SALDO"
    })

    # Limpiar SALDO (Regla B1)
    def clean_saldo(x):
        if pd.isna(x):
            return None
        x = str(x).replace("$", "").replace(",", "").strip()
        if x in ["-", "", " -", "- ", "$-"]:
            return None
        try:
            return float(x)
        except:
            return None

    df["SALDO"] = df["SALDO"].apply(clean_saldo)
    df = df[df["SALDO"].notna()]

    # Convertir FECHAVTO a fecha
    df["FECHAVTO"] = pd.to_datetime(df["FECHAVTO"], format="%d/%m/%Y")

    # Meses en español
    meses_es = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }

    df["MES"] = df["FECHAVTO"].dt.month.map(meses_es)
    df["AÑO"] = df["FECHAVTO"].dt.year

    return df

# -----------------------------
# 2. Generar XLSX
# -----------------------------
def generate_xlsx(df: pd.DataFrame, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte CxP"

    bold = Font(bold=True)
    bold_big = Font(bold=True, size=14)
    right = Alignment(horizontal="right")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    row = 1

    # Proveedores ordenados alfabéticamente
    for proveedor in sorted(df["PROVEEDOR"].unique()):
        sub = df[df["PROVEEDOR"] == proveedor]

        # Título del proveedor
        ws.cell(row=row, column=1, value=proveedor).font = bold_big
        row += 2

        # Encabezados
        headers = ["FECHAVTO", "NOFAC", "SALDO"]
        for col, h in enumerate(headers, start=1):
            c = ws.cell(row=row, column=col, value=h)
            c.font = bold
            c.border = border
        row += 1

        # Agrupar por mes y año
        for (mes, año), grupo in sub.groupby(["MES", "AÑO"]):
            for _, r in grupo.iterrows():
                # FECHAVTO
                c1 = ws.cell(row=row, column=1, value=r["FECHAVTO"].strftime("%d/%m/%Y"))
                c1.border = border

                # NOFAC
                c2 = ws.cell(row=row, column=2, value=r["NOFAC"])
                c2.border = border

                # SALDO
                c3 = ws.cell(row=row, column=3, value=r["SALDO"])
                c3.alignment = right
                c3.border = border

                row += 1

            # Fila de total mensual
            total_mes = grupo["SALDO"].sum()
            c_tot_label = ws.cell(row=row, column=1, value=f"Total a pagar {mes} {año}")
            c_tot_label.font = bold

            c_tot_val = ws.cell(row=row, column=3, value=total_mes)
            c_tot_val.font = bold
            c_tot_val.alignment = right

            row += 2

        # Espacio entre proveedores
        row += 1

    # Consolidado general por mes
    ws.cell(row=row, column=1, value="CONSOLIDADO GENERAL POR MES").font = bold_big
    row += 2

    consolidado = df.groupby(["MES", "AÑO"])["SALDO"].sum().reset_index()

    for _, r in consolidado.iterrows():
        c1 = ws.cell(row=row, column=1, value=f"{r['MES']} {r['AÑO']}")
        c1.font = bold

        c2 = ws.cell(row=row, column=2, value=r["SALDO"])
        c2.font = bold
        c2.alignment = right

        row += 1

    # Ajustar anchos de columna
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 22

    wb.save(output_path)

# -----------------------------
# 3. Punto de entrada
# -----------------------------
def run_report(input_path: str, base_name: str = "ReporteCXP", output_dir: str = "output"):
    os.makedirs(output_dir, exist_ok=True)
    today = datetime.today().strftime("%Y%m%d")
    xlsx_name = f"{base_name}{today}.xlsx"
    output_path = os.path.join(output_dir, xlsx_name)

    df = load_and_clean(input_path)
    generate_xlsx(df, output_path)

    print(f"Archivo generado: {output_path}")