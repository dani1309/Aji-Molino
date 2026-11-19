import base64
import datetime
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import streamlit as st

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA Y RUTAS
# ==============================================================================
st.set_page_config(
    page_title="Ají y Molino — Inventario",
    page_icon="🌶️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "LISTADO_VENTA.xlsx"
BACKUP_DIR = BASE_DIR / "Backups"
GEMINI_CONFIG_PATH = BASE_DIR / "gemini_config.json"


def create_daily_backup(max_backups: int = 15):
    """Crea una copia de seguridad diaria en Backups/ antes de modificar el Excel."""
    if not EXCEL_PATH.exists():
        return
    BACKUP_DIR.mkdir(exist_ok=True)
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    backup_file = BACKUP_DIR / f"LISTADO_VENTA_{today_str}.xlsx"
    if not backup_file.exists():
        shutil.copyfile(EXCEL_PATH, backup_file)
        # Conservar solo los últimos max_backups archivos
        backups = sorted(BACKUP_DIR.glob("LISTADO_VENTA_*.xlsx"))
        for old_b in backups[:-max_backups]:
            try:
                old_b.unlink()
            except OSError:
                pass

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] {
        background-color: var(--secondary-background-color, rgba(128, 128, 128, 0.12));
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-left: 4px solid #C0392B;
        padding: 12px 16px;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# LECTURA Y ESCRITURA EN LISTADO_VENTA.xlsx
# ==============================================================================
def to_float(val, default=0.0):
    if val is None or val == "":
        return default
    try:
        return float(val)
    except Exception:
        return default


def to_date(val):
    if val is None or val == "":
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def restore_data_validations(wb):
    """Mantiene las validaciones de datos intactas al guardar con openpyxl."""
    ws_inv = wb["📦 INVENTARIO"]
    ws_mov = wb["🔄 MOVIMIENTOS"]
    ws_dash = wb["📊 DASHBOARD"]
    last_inv_row = ws_inv.max_row

    ws_mov.data_validations.dataValidation.clear()
    dv_fecha_mov = DataValidation(
        type="date", operator="between", formula1="2020-01-01", formula2="2035-12-31", allow_blank=True
    )
    ws_mov.add_data_validation(dv_fecha_mov)
    dv_fecha_mov.add("A12:A311")

    dv_tipo = DataValidation(type="list", formula1='"SALIDA,ENTRADA,AJUSTE"', allow_blank=True)
    ws_mov.add_data_validation(dv_tipo)
    dv_tipo.add("B12:B311")

    dv_prod = DataValidation(
        type="list", formula1=f"='📦 INVENTARIO'!$AA$6:$AA${last_inv_row}", allow_blank=True
    )
    ws_mov.add_data_validation(dv_prod)
    dv_prod.add("C12:C311")

    dv_pago = DataValidation(
        type="list",
        formula1='"Efectivo,Nequi / Daviplata,Transferencia Bancaria,Tarjeta / Datáfono,Crédito Cliente,Factura Proveedor,Ajuste Interno"',
        allow_blank=True,
    )
    ws_mov.add_data_validation(dv_pago)
    dv_pago.add("K12:K311")

    ws_dash.data_validations.dataValidation.clear()
    dv_fecha_dash = DataValidation(
        type="date", operator="between", formula1="2020-01-01", formula2="2035-12-31", allow_blank=True
    )
    ws_dash.add_data_validation(dv_fecha_dash)
    dv_fecha_dash.add("C32")
    dv_fecha_dash.add("E32")

    dv_dash_prod = DataValidation(
        type="list", formula1=f"='📦 INVENTARIO'!$C$6:$C${last_inv_row}", allow_blank=True
    )
    ws_dash.add_data_validation(dv_dash_prod)
    dv_dash_prod.add("G32")


def is_cloud_mode():
    """Retorna True si la app está configurada con Google Sheets en st.secrets."""
    try:
        return "gcp_service_account" in st.secrets and ("spreadsheet_id" in st.secrets or "spreadsheet_url" in st.secrets)
    except Exception:
        return False


def _get_gsheets_client():
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
    gc = gspread.authorize(creds)
    if "spreadsheet_id" in st.secrets:
        return gc.open_by_key(st.secrets["spreadsheet_id"])
    return gc.open_by_url(st.secrets["spreadsheet_url"])


def _compute_inventory_dfs(inv_rows, mov_rows):
    df_inv = pd.DataFrame(inv_rows)
    df_mov = pd.DataFrame(mov_rows)

    if not df_mov.empty:
        valid_mov = df_mov[df_mov["Tipo"].isin(["ENTRADA", "SALIDA", "AJUSTE"])]
        ent_sum = valid_mov[valid_mov["Tipo"] == "ENTRADA"].groupby("Producto")["Cantidad"].sum()
        sal_sum = valid_mov[valid_mov["Tipo"] == "SALIDA"].groupby("Producto")["Cantidad"].sum()
        aju_sum = valid_mov[valid_mov["Tipo"] == "AJUSTE"].groupby("Producto")["Cantidad"].sum()
    else:
        ent_sum = pd.Series(dtype=int)
        sal_sum = pd.Series(dtype=int)
        aju_sum = pd.Series(dtype=int)

    df_inv["Entradas (+)"] = df_inv["Producto"].map(ent_sum).fillna(0).astype(int)
    df_inv["Salidas (-)"] = df_inv["Producto"].map(sal_sum).fillna(0).astype(int)
    df_inv["Ajustes (+/-)"] = df_inv["Producto"].map(aju_sum).fillna(0).astype(int)
    df_inv["Stock Actual"] = (
        df_inv["Stock Inicial"] + df_inv["Entradas (+)"] - df_inv["Salidas (-)"] + df_inv["Ajustes (+/-)"]
    )

    def calc_estado(row):
        if row["Stock Actual"] <= 0:
            return "🔴 AGOTADO"
        if row["Stock Actual"] <= row["Stock Mín."]:
            return "🟡 PEDIR"
        return "🟢 OK"

    df_inv["Estado"] = df_inv.apply(calc_estado, axis=1)
    df_inv["Valor Inv. Costo"] = df_inv["Stock Actual"] * df_inv["Costo c/IVA"]
    df_inv["Valor Inv. Venta"] = df_inv["Stock Actual"] * df_inv["PVF"]
    df_inv["Ganancia x Und"] = df_inv.apply(
        lambda r: (r["PVF"] - r["Costo c/IVA"]) if (r["PVF"] > 0 and r["Costo c/IVA"] > 0) else 0, axis=1
    )
    return df_inv, df_mov


@st.cache_data(ttl=30, show_spinner=False)
def _load_gsheets_cached():
    sh = _get_gsheets_client()
    ws_inv = sh.worksheet("📦 INVENTARIO")
    ws_mov = sh.worksheet("🔄 MOVIMIENTOS")

    inv_vals = ws_inv.get_all_values()
    mov_vals = ws_mov.get_all_values()

    def cell_at(row_list, col_1idx):
        return row_list[col_1idx - 1] if len(row_list) >= col_1idx else None

    def parse_cop_str(v, default=0.0):
        if v is None or v == "":
            return default
        s = str(v).replace("$", "").replace(",", "").replace("%", "").strip()
        try:
            num = float(s)
            return num / 100.0 if "%" in str(v) else num
        except Exception:
            return default

    inv_rows = []
    for idx_r, row_l in enumerate(inv_vals[5:], start=6):
        sku = cell_at(row_l, 1)
        desc = cell_at(row_l, 3)
        if not sku and not desc:
            continue
        stock_ini = int(round(parse_cop_str(cell_at(row_l, 6), 0)))
        stock_min = int(round(parse_cop_str(cell_at(row_l, 11), 0)))
        costo_base = round(parse_cop_str(cell_at(row_l, 13), 0))
        iva = parse_cop_str(cell_at(row_l, 14), 0.0)
        if iva > 1:
            iva = iva / 100.0
        costo_iva = round(costo_base * (1 + iva))
        margen_sug = parse_cop_str(cell_at(row_l, 16), 0.30)
        if margen_sug > 1:
            margen_sug = margen_sug / 100.0
        pvf = round(parse_cop_str(cell_at(row_l, 18), 0))
        p_alt_raw = cell_at(row_l, 19)
        precio_alt = round(parse_cop_str(p_alt_raw, 0)) if p_alt_raw not in (None, "", "-") else None
        inv_rows.append(
            {
                "fila_excel": idx_r,
                "SKU": str(sku or "").strip(),
                "Categoría": str(cell_at(row_l, 2) or "").strip(),
                "Producto": str(desc or "").strip(),
                "Presentación": str(cell_at(row_l, 4) or "").strip(),
                "Origen": str(cell_at(row_l, 5) or "").strip(),
                "Stock Inicial": stock_ini,
                "Stock Mín.": stock_min,
                "Costo Base": costo_base,
                "IVA %": iva,
                "Costo c/IVA": costo_iva,
                "Margen Sug.": margen_sug,
                "PVF": pvf,
                "Precio Alt.": precio_alt,
                "Observaciones": str(cell_at(row_l, 24) or ""),
            }
        )

    prod_map = {r["Producto"]: r for r in inv_rows}
    mov_rows = []
    for idx_r, row_l in enumerate(mov_vals[11:], start=12):
        fecha_raw = cell_at(row_l, 1)
        tipo = str(cell_at(row_l, 2) or "").strip().upper()
        prod = str(cell_at(row_l, 3) or "").strip()
        qty_raw = cell_at(row_l, 7)
        if not prod and not tipo and not qty_raw and not fecha_raw:
            continue
        fecha_d = to_date(fecha_raw)
        qty = int(round(parse_cop_str(qty_raw, 0)))
        info = prod_map.get(prod, {})
        v_unit = info.get("PVF", 0) if tipo == "SALIDA" else info.get("Costo c/IVA", 0)
        lote_id = f"MOV-{idx_r - 11:03d}" + (f"-{fecha_d.strftime('%d%m')}" if fecha_d else "")
        mov_rows.append(
            {
                "fila_excel": idx_r,
                "ID Lote": lote_id,
                "Fecha": fecha_d,
                "Tipo": tipo if tipo else "⚠️ SIN TIPO",
                "Producto": prod,
                "SKU": info.get("SKU", ""),
                "Categoría": info.get("Categoría", ""),
                "Cantidad": qty,
                "Valor Unitario": v_unit,
                "Total": qty * v_unit,
                "Cliente / Proveedor": str(cell_at(row_l, 10) or "").strip(),
                "Medio de Pago": str(cell_at(row_l, 11) or "").strip(),
                "Observaciones": str(cell_at(row_l, 12) or "").strip(),
            }
        )

    return _compute_inventory_dfs(inv_rows, mov_rows)


def _clear_all_caches():
    _load_data_cached.clear()
    _load_gsheets_cached.clear()


@st.cache_data(show_spinner=False)
def _load_data_cached(file_mtime: float):
    # Leer desde copia temporal para que funcione incluso si el usuario tiene el Excel abierto
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)
    try:
        try:
            shutil.copyfile(EXCEL_PATH, tmp_path)
        except PermissionError:
            # PowerShell Copy-Item respeta el bloqueo compartido de Excel/OneDrive
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"Copy-Item -LiteralPath '{EXCEL_PATH}' -Destination '{tmp_path}' -Force",
                ],
                check=True,
                capture_output=True,
            )
        wb = openpyxl.load_workbook(tmp_path, data_only=False)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    ws_inv = wb["📦 INVENTARIO"]
    ws_mov = wb["🔄 MOVIMIENTOS"]

    # 1. Leer productos base desde 📦 INVENTARIO
    inv_rows = []
    for r in range(6, ws_inv.max_row + 1):
        sku = ws_inv.cell(r, 1).value
        desc = ws_inv.cell(r, 3).value
        if not sku and not desc:
            continue
        stock_ini = int(round(to_float(ws_inv.cell(r, 6).value, 0)))
        stock_min = int(round(to_float(ws_inv.cell(r, 11).value, 0)))
        costo_base = round(to_float(ws_inv.cell(r, 13).value, 0))
        iva = to_float(ws_inv.cell(r, 14).value, 0.0)
        costo_iva = round(costo_base * (1 + iva))
        margen_sug = to_float(ws_inv.cell(r, 16).value, 0.30)
        pvf = round(to_float(ws_inv.cell(r, 18).value, 0))
        precio_alt = ws_inv.cell(r, 19).value
        precio_alt = round(to_float(precio_alt, 0)) if precio_alt not in (None, "") else None
        obs = str(ws_inv.cell(r, 24).value or "")

        inv_rows.append(
            {
                "fila_excel": r,
                "SKU": str(sku or "").strip(),
                "Categoría": str(ws_inv.cell(r, 2).value or "").strip(),
                "Producto": str(desc or "").strip(),
                "Presentación": str(ws_inv.cell(r, 4).value or "").strip(),
                "Origen": str(ws_inv.cell(r, 5).value or "").strip(),
                "Stock Inicial": stock_ini,
                "Stock Mín.": stock_min,
                "Costo Base": costo_base,
                "IVA %": iva,
                "Costo c/IVA": costo_iva,
                "Margen Sug.": margen_sug,
                "PVF": pvf,
                "Precio Alt.": precio_alt,
                "Observaciones": obs,
            }
        )

    df_inv = pd.DataFrame(inv_rows)
    prod_map = df_inv.set_index("Producto").to_dict(orient="index")

    # 2. Leer movimientos desde 🔄 MOVIMIENTOS (filas 12..500)
    mov_rows = []
    for r in range(12, ws_mov.max_row + 1):
        fecha_raw = ws_mov.cell(r, 1).value
        tipo = str(ws_mov.cell(r, 2).value or "").strip().upper()
        prod = str(ws_mov.cell(r, 3).value or "").strip()
        qty_raw = ws_mov.cell(r, 7).value
        cli = str(ws_mov.cell(r, 10).value or "").strip()
        pago = str(ws_mov.cell(r, 11).value or "").strip()
        obs = str(ws_mov.cell(r, 12).value or "").strip()

        if not prod and not tipo and qty_raw is None and fecha_raw is None:
            continue

        fecha_d = to_date(fecha_raw)
        qty = int(round(to_float(qty_raw, 0)))
        info = prod_map.get(prod, {})
        sku = info.get("SKU", "")
        cat = info.get("Categoría", "")
        v_unit = info.get("PVF", 0) if tipo == "SALIDA" else info.get("Costo c/IVA", 0)
        total = qty * v_unit
        lote_id = f"MOV-{r - 11:03d}" + (f"-{fecha_d.strftime('%d%m')}" if fecha_d else "")

        mov_rows.append(
            {
                "fila_excel": r,
                "ID Lote": lote_id,
                "Fecha": fecha_d,
                "Tipo": tipo if tipo else "⚠️ SIN TIPO",
                "Producto": prod,
                "SKU": sku,
                "Categoría": cat,
                "Cantidad": qty,
                "Valor Unitario": v_unit,
                "Total": total,
                "Cliente / Proveedor": cli,
                "Medio de Pago": pago,
                "Observaciones": obs,
            }
        )

    df_mov = pd.DataFrame(mov_rows)

    # 3. Calcular Stock Actual, Entradas, Salidas, Ajustes y Estado en vivo
    if not df_mov.empty:
        valid_mov = df_mov[df_mov["Tipo"].isin(["ENTRADA", "SALIDA", "AJUSTE"])]
        ent_sum = valid_mov[valid_mov["Tipo"] == "ENTRADA"].groupby("Producto")["Cantidad"].sum()
        sal_sum = valid_mov[valid_mov["Tipo"] == "SALIDA"].groupby("Producto")["Cantidad"].sum()
        aju_sum = valid_mov[valid_mov["Tipo"] == "AJUSTE"].groupby("Producto")["Cantidad"].sum()
    else:
        ent_sum = pd.Series(dtype=int)
        sal_sum = pd.Series(dtype=int)
        aju_sum = pd.Series(dtype=int)

    df_inv["Entradas (+)"] = df_inv["Producto"].map(ent_sum).fillna(0).astype(int)
    df_inv["Salidas (-)"] = df_inv["Producto"].map(sal_sum).fillna(0).astype(int)
    df_inv["Ajustes (+/-)"] = df_inv["Producto"].map(aju_sum).fillna(0).astype(int)
    df_inv["Stock Actual"] = (
        df_inv["Stock Inicial"] + df_inv["Entradas (+)"] - df_inv["Salidas (-)"] + df_inv["Ajustes (+/-)"]
    )

    def calc_estado(row):
        if row["Stock Actual"] <= 0:
            return "🔴 AGOTADO"
        if row["Stock Actual"] <= row["Stock Mín."]:
            return "🟡 PEDIR"
        return "🟢 OK"

    df_inv["Estado"] = df_inv.apply(calc_estado, axis=1)
    df_inv["Valor Inv. Costo"] = df_inv["Stock Actual"] * df_inv["Costo c/IVA"]
    df_inv["Valor Inv. Venta"] = df_inv["Stock Actual"] * df_inv["PVF"]
    df_inv["Ganancia x Und"] = df_inv.apply(
        lambda r: (r["PVF"] - r["Costo c/IVA"]) if (r["PVF"] > 0 and r["Costo c/IVA"] > 0) else 0, axis=1
    )

    return df_inv, df_mov


def load_data():
    if is_cloud_mode():
        try:
            return _load_gsheets_cached()
        except Exception as e:
            st.error(f"❌ Error conectando a Google Sheets: {e}")
            st.stop()

    if not EXCEL_PATH.exists():
        st.error(f"No se encontró el archivo: {EXCEL_PATH}")
        st.stop()
    try:
        mtime = EXCEL_PATH.stat().st_mtime
        return _load_data_cached(mtime)
    except Exception as e:
        st.error(
            f"❌ No se pudo leer `{EXCEL_PATH.name}` porque está bloqueado por Excel/OneDrive. "
            f"Ciérralo en Excel y recarga esta página. ({e})"
        )
        st.stop()


def _gs_write_movement_row(ws_mov, row_num, fecha_val, tipo, producto, cantidad, cliente, pago, obs):
    fecha_str = fecha_val.strftime("%d/%m/%Y") if fecha_val else ""
    ws_mov.update_cell(row_num, 1, fecha_str)
    ws_mov.update_cell(row_num, 2, tipo)
    ws_mov.update_cell(row_num, 3, producto)
    ws_mov.update_cell(row_num, 7, int(cantidad) if cantidad else "")
    ws_mov.update_cell(row_num, 10, cliente or "")
    ws_mov.update_cell(row_num, 11, pago or "")
    ws_mov.update_cell(row_num, 12, obs or "")


def save_movement(fecha_val, tipo, producto, cantidad, cliente, pago, obs):
    if is_cloud_mode():
        rows = save_cart_movements([
            {
                "fecha": fecha_val,
                "tipo": tipo,
                "producto": producto,
                "cantidad": cantidad,
                "cliente": cliente,
                "pago": pago,
                "obs": obs,
            }
        ])
        return rows[0]

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_mov = wb["🔄 MOVIMIENTOS"]

    # Buscar primera fila vacía en 🔄 MOVIMIENTOS (desde la 12)
    target_row = None
    for r in range(12, 501):
        if (
            ws_mov.cell(r, 1).value is None
            and ws_mov.cell(r, 2).value is None
            and ws_mov.cell(r, 3).value is None
            and ws_mov.cell(r, 7).value is None
        ):
            target_row = r
            break

    if target_row is None:
        raise ValueError("Se alcanzó el límite de 500 filas en 🔄 MOVIMIENTOS.")

    ws_mov.cell(target_row, 1, fecha_val).number_format = "dd/mm/yyyy"
    ws_mov.cell(target_row, 2, tipo)
    ws_mov.cell(target_row, 3, producto)
    ws_mov.cell(target_row, 7, int(cantidad))
    ws_mov.cell(target_row, 10, cliente)
    ws_mov.cell(target_row, 11, pago)
    ws_mov.cell(target_row, 12, obs)

    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _load_data_cached.clear()


def update_inventory_bulk(changes: list[dict]):
    """Actualiza múltiples filas de 📦 INVENTARIO (Stock Inicial, Stock Mín., Costo Base, PVF, Observaciones)."""
    if not changes:
        return 0
    if is_cloud_mode():
        sh = _get_gsheets_client()
        ws_inv = sh.worksheet("📦 INVENTARIO")
        updates = []
        for ch in changes:
            r = int(ch["fila_excel"])
            updates.extend([
                {"range": f"F{r}", "values": [[int(ch["stock_ini"])]]},
                {"range": f"K{r}", "values": [[int(ch["stock_min"])]]},
                {"range": f"M{r}", "values": [[int(ch["costo_base"])]]},
                {"range": f"R{r}", "values": [[int(ch["pvf"])]]},
                {"range": f"X{r}", "values": [[str(ch["obs"] or "")]]},
            ])
        ws_inv.batch_update(updates, value_input_option="USER_ENTERED")
        _clear_all_caches()
        return len(changes)

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_inv = wb["📦 INVENTARIO"]
    for ch in changes:
        r = int(ch["fila_excel"])
        ws_inv.cell(r, 6, int(ch["stock_ini"]))
        ws_inv.cell(r, 11, int(ch["stock_min"]))
        ws_inv.cell(r, 13, int(ch["costo_base"]))
        ws_inv.cell(r, 18, int(ch["pvf"]))
        ws_inv.cell(r, 24, str(ch["obs"] or ""))
    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _clear_all_caches()
    return len(changes)


def update_product_in_excel(fila_excel, stock_ini, stock_min, costo_base, pvf, obs):
    if is_cloud_mode():
        sh = _get_gsheets_client()
        ws_inv = sh.worksheet("📦 INVENTARIO")
        ws_inv.update_cell(fila_excel, 6, int(stock_ini))
        ws_inv.update_cell(fila_excel, 11, int(stock_min))
        ws_inv.update_cell(fila_excel, 13, int(costo_base))
        ws_inv.update_cell(fila_excel, 18, int(pvf))
        ws_inv.update_cell(fila_excel, 24, obs)
        _clear_all_caches()
        return

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_inv = wb["📦 INVENTARIO"]
    ws_inv.cell(fila_excel, 6, int(stock_ini))
    ws_inv.cell(fila_excel, 11, int(stock_min))
    ws_inv.cell(fila_excel, 13, int(costo_base))
    ws_inv.cell(fila_excel, 18, int(pvf))
    ws_inv.cell(fila_excel, 24, obs)
    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _load_data_cached.clear()


def delete_movement_row(fila_excel):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_mov = wb["🔄 MOVIMIENTOS"]
    for col in (1, 2, 3, 7, 10, 11, 12):
        ws_mov.cell(fila_excel, col, None)
    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _load_data_cached.clear()


def save_movements_table(edited_df):
    """Sincroniza toda la tabla editada en Streamlit hacia 🔄 MOVIMIENTOS en Excel o Google Sheets."""
    if is_cloud_mode():
        sh = _get_gsheets_client()
        ws_mov = sh.worksheet("🔄 MOVIMIENTOS")
        all_vals = ws_mov.get_all_values()
        max_r = max(len(all_vals), 12)
        # Limpiar columnas editables A,B,C,G,J,K,L desde fila 12
        updates = []
        for r in range(12, max_r + 1):
            for col_l in ("A", "B", "C", "G", "J", "K", "L"):
                updates.append({"range": f"{col_l}{r}", "values": [[""]]})
        row_idx = 12
        for _, row in edited_df.iterrows():
            prod = str(row.get("Producto") or "").strip()
            if not prod or prod.lower() in ("nan", "none"):
                continue
            fecha_v = to_date(row.get("Fecha"))
            tipo_v = str(row.get("Tipo") or "").strip().upper()
            if tipo_v in ("⚠️ SIN TIPO", "NAN", "NONE"):
                tipo_v = ""
            qty_v = int(round(to_float(row.get("Cantidad"), 0)))
            cli_v = "" if pd.isna(row.get("Cliente / Proveedor")) else str(row.get("Cliente / Proveedor") or "").strip()
            pago_v = "" if pd.isna(row.get("Medio de Pago")) else str(row.get("Medio de Pago") or "").strip()
            obs_v = "" if pd.isna(row.get("Observaciones")) else str(row.get("Observaciones") or "").strip()
            updates.extend([
                {"range": f"A{row_idx}", "values": [[fecha_v.strftime("%d/%m/%Y") if fecha_v else ""]]},
                {"range": f"B{row_idx}", "values": [[tipo_v]]},
                {"range": f"C{row_idx}", "values": [[prod]]},
                {"range": f"G{row_idx}", "values": [[qty_v if qty_v != 0 else ""]]},
                {"range": f"J{row_idx}", "values": [[cli_v]]},
                {"range": f"K{row_idx}", "values": [[pago_v]]},
                {"range": f"L{row_idx}", "values": [[obs_v]]},
            ])
            row_idx += 1
        if updates:
            ws_mov.batch_update(updates, value_input_option="USER_ENTERED")
        _clear_all_caches()
        return

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_mov = wb["🔄 MOVIMIENTOS"]

    # Limpiar filas previas (12..500) en columnas editables
    for r in range(12, 501):
        has_any = any(ws_mov.cell(r, c).value is not None for c in (1, 2, 3, 7, 10, 11, 12))
        if has_any:
            for c in (1, 2, 3, 7, 10, 11, 12):
                ws_mov.cell(r, c, None)

    row_idx = 12
    for _, row in edited_df.iterrows():
        prod = str(row.get("Producto") or "").strip()
        if not prod or prod.lower() in ("nan", "none"):
            continue
        fecha_v = to_date(row.get("Fecha"))
        tipo_v = str(row.get("Tipo") or "").strip().upper()
        if tipo_v in ("⚠️ SIN TIPO", "NAN", "NONE"):
            tipo_v = ""
        qty_v = int(round(to_float(row.get("Cantidad"), 0)))
        cli_v = "" if pd.isna(row.get("Cliente / Proveedor")) else str(row.get("Cliente / Proveedor") or "").strip()
        pago_v = "" if pd.isna(row.get("Medio de Pago")) else str(row.get("Medio de Pago") or "").strip()
        obs_v = "" if pd.isna(row.get("Observaciones")) else str(row.get("Observaciones") or "").strip()

        if fecha_v:
            ws_mov.cell(row_idx, 1, fecha_v).number_format = "dd/mm/yyyy"
        ws_mov.cell(row_idx, 2, tipo_v if tipo_v else None)
        ws_mov.cell(row_idx, 3, prod)
        ws_mov.cell(row_idx, 7, qty_v if qty_v != 0 else None)
        ws_mov.cell(row_idx, 10, cli_v if cli_v else None)
        ws_mov.cell(row_idx, 11, pago_v if pago_v else None)
        ws_mov.cell(row_idx, 12, obs_v if obs_v else None)
        row_idx += 1

    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _load_data_cached.clear()


def save_cart_movements(cart_items):
    """Guarda múltiples movimientos del carrito en 🔄 MOVIMIENTOS de una sola vez."""
    if is_cloud_mode():
        sh = _get_gsheets_client()
        ws_mov = sh.worksheet("🔄 MOVIMIENTOS")
        all_vals = ws_mov.get_all_values()
        # Buscar primera fila vacía desde la 12
        start_search = 12
        saved_rows = []
        updates = []
        for item in cart_items:
            target_row = None
            for r in range(start_search, 501):
                row_l = all_vals[r - 1] if r - 1 < len(all_vals) else []
                c_a = row_l[0] if len(row_l) > 0 else ""
                c_b = row_l[1] if len(row_l) > 1 else ""
                c_c = row_l[2] if len(row_l) > 2 else ""
                c_g = row_l[6] if len(row_l) > 6 else ""
                if not c_a and not c_b and not c_c and not c_g:
                    target_row = r
                    start_search = r + 1
                    break
            if target_row is None:
                raise ValueError("Se alcanzó el límite de 500 filas en 🔄 MOVIMIENTOS.")
            f_str = item["fecha"].strftime("%d/%m/%Y") if item["fecha"] else ""
            updates.extend([
                {"range": f"A{target_row}", "values": [[f_str]]},
                {"range": f"B{target_row}", "values": [[item["tipo"]]]},
                {"range": f"C{target_row}", "values": [[item["producto"]]]},
                {"range": f"G{target_row}", "values": [[int(item["cantidad"])]]},
                {"range": f"J{target_row}", "values": [[item["cliente"] or ""]]},
                {"range": f"K{target_row}", "values": [[item["pago"] or ""]]},
                {"range": f"L{target_row}", "values": [[item["obs"] or ""]]},
            ])
            saved_rows.append(target_row)
        if updates:
            ws_mov.batch_update(updates, value_input_option="USER_ENTERED")
        _clear_all_caches()
        return saved_rows

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_mov = wb["🔄 MOVIMIENTOS"]

    saved_rows = []
    start_search = 12
    for item in cart_items:
        target_row = None
        for r in range(start_search, 501):
            if (
                ws_mov.cell(r, 1).value is None
                and ws_mov.cell(r, 2).value is None
                and ws_mov.cell(r, 3).value is None
                and ws_mov.cell(r, 7).value is None
            ):
                target_row = r
                start_search = r + 1
                break
        if target_row is None:
            raise ValueError("Se alcanzó el límite de 500 filas en 🔄 MOVIMIENTOS.")

        ws_mov.cell(target_row, 1, item["fecha"]).number_format = "dd/mm/yyyy"
        ws_mov.cell(target_row, 2, item["tipo"])
        ws_mov.cell(target_row, 3, item["producto"])
        ws_mov.cell(target_row, 7, int(item["cantidad"]))
        ws_mov.cell(target_row, 10, item["cliente"] or None)
        ws_mov.cell(target_row, 11, item["pago"] or None)
        ws_mov.cell(target_row, 12, item["obs"] or None)
        saved_rows.append(target_row)

    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _load_data_cached.clear()


def create_new_product_in_excel(sku, categoria, descripcion, presentacion, stock_ini, stock_min, costo_base, iva_pct, margen_sug, pvf, obs):
    """Agrega un producto nuevo al final de 📦 INVENTARIO con todas sus fórmulas."""
    if is_cloud_mode():
        sh = _get_gsheets_client()
        ws_inv = sh.worksheet("📦 INVENTARIO")
        r = len(ws_inv.get_all_values()) + 1
        idx_num = r - 5
        row_vals = [
            sku.strip(),
            categoria.strip(),
            descripcion.strip(),
            presentacion.strip() or "und",
            "App Streamlit",
            int(stock_ini),
            f'=SUMIFS(\'🔄 MOVIMIENTOS\'!$G$12:$G$500, \'🔄 MOVIMIENTOS\'!$C$12:$C$500, C{r}, \'🔄 MOVIMIENTOS\'!$B$12:$B$500, "ENTRADA")',
            f'=SUMIFS(\'🔄 MOVIMIENTOS\'!$G$12:$G$500, \'🔄 MOVIMIENTOS\'!$C$12:$C$500, C{r}, \'🔄 MOVIMIENTOS\'!$B$12:$B$500, "SALIDA")',
            f'=SUMIFS(\'🔄 MOVIMIENTOS\'!$G$12:$G$500, \'🔄 MOVIMIENTOS\'!$C$12:$C$500, C{r}, \'🔄 MOVIMIENTOS\'!$B$12:$B$500, "AJUSTE")',
            f"=F{r}+G{r}-H{r}+I{r}",
            int(stock_min),
            f'=IF(J{r}<=0, "🔴 AGOTADO", IF(J{r}<=K{r}, "🟡 PEDIR", "🟢 OK"))',
            int(costo_base),
            float(iva_pct),
            f"=ROUND(M{r}*(1+N{r}), 0)",
            float(margen_sug),
            f"=IF(O{r}>0, ROUND(O{r}/(1-P{r}), 0), 0)",
            int(pvf),
            "",
            f"=IF(AND(R{r}>0, O{r}>0), R{r}-O{r}, 0)",
            f"=IF(AND(R{r}>0, O{r}>0), (R{r}-O{r})/R{r}, 0)",
            f"=J{r}*O{r}",
            f"=J{r}*R{r}",
            obs.strip(),
            "",
            f'=IF(ISNUMBER(SEARCH(\'🔄 MOVIMIENTOS\'!$C$4, C{r}&" "&A{r}&" "&B{r})), MAX($Z$5:Z{r-1})+1, 0)',
            f'=IFERROR(INDEX($C$6:$C${r}, MATCH({idx_num}, $Z$6:$Z${r}, 0)), "")',
        ]
        ws_inv.update(f"A{r}:AA{r}", [row_vals], value_input_option="USER_ENTERED")
        _clear_all_caches()
        return r

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_inv = wb["📦 INVENTARIO"]
    r = ws_inv.max_row + 1
    idx_num = r - 5

    ws_inv.cell(r, 1, sku.strip())
    ws_inv.cell(r, 2, categoria.strip())
    ws_inv.cell(r, 3, descripcion.strip())
    ws_inv.cell(r, 4, presentacion.strip() or "und")
    ws_inv.cell(r, 5, "App Streamlit")
    ws_inv.cell(r, 6, int(stock_ini))
    ws_inv.cell(r, 7, f'=SUMIFS(\'🔄 MOVIMIENTOS\'!$G$12:$G$500, \'🔄 MOVIMIENTOS\'!$C$12:$C$500, C{r}, \'🔄 MOVIMIENTOS\'!$B$12:$B$500, "ENTRADA")')
    ws_inv.cell(r, 8, f'=SUMIFS(\'🔄 MOVIMIENTOS\'!$G$12:$G$500, \'🔄 MOVIMIENTOS\'!$C$12:$C$500, C{r}, \'🔄 MOVIMIENTOS\'!$B$12:$B$500, "SALIDA")')
    ws_inv.cell(r, 9, f'=SUMIFS(\'🔄 MOVIMIENTOS\'!$G$12:$G$500, \'🔄 MOVIMIENTOS\'!$C$12:$C$500, C{r}, \'🔄 MOVIMIENTOS\'!$B$12:$B$500, "AJUSTE")')
    ws_inv.cell(r, 10, f"=F{r}+G{r}-H{r}+I{r}")
    ws_inv.cell(r, 11, int(stock_min))
    ws_inv.cell(r, 12, f'=IF(J{r}<=0, "🔴 AGOTADO", IF(J{r}<=K{r}, "🟡 PEDIR", "🟢 OK"))')
    ws_inv.cell(r, 13, int(costo_base))
    ws_inv.cell(r, 14, float(iva_pct))
    ws_inv.cell(r, 15, f"=ROUND(M{r}*(1+N{r}), 0)")
    ws_inv.cell(r, 16, float(margen_sug))
    ws_inv.cell(r, 17, f"=IF(O{r}>0, ROUND(O{r}/(1-P{r}), 0), 0)")
    ws_inv.cell(r, 18, int(pvf))
    ws_inv.cell(r, 19, None)
    ws_inv.cell(r, 20, f"=IF(AND(R{r}>0, O{r}>0), R{r}-O{r}, 0)")
    ws_inv.cell(r, 21, f"=IF(AND(R{r}>0, O{r}>0), (R{r}-O{r})/R{r}, 0)")
    ws_inv.cell(r, 22, f"=J{r}*O{r}")
    ws_inv.cell(r, 23, f"=J{r}*R{r}")
    ws_inv.cell(r, 24, obs.strip())
    ws_inv.cell(r, 26, f'=IF(ISNUMBER(SEARCH(\'🔄 MOVIMIENTOS\'!$C$4, C{r}&" "&A{r}&" "&B{r})), MAX($Z$5:Z{r-1})+1, 0)')
    ws_inv.cell(r, 27, f'=IFERROR(INDEX($C$6:$C${r}, MATCH({idx_num}, $Z$6:$Z${r}, 0)), "")')

    ws_inv.auto_filter.ref = f"A5:X{r}"
    restore_data_validations(wb)
    create_daily_backup()
    wb.save(EXCEL_PATH)
    _load_data_cached.clear()


def get_gemini_api_key():
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        pass
    if GEMINI_CONFIG_PATH.exists():
        try:
            cfg = json.loads(GEMINI_CONFIG_PATH.read_text(encoding="utf-8"))
            return cfg.get("api_key", "").strip()
        except Exception:
            return ""
    return os.environ.get("GEMINI_API_KEY", "").strip()


def parse_notebook_image_with_gemini(image_bytes: bytes, mime_type: str, catalog_names: list[str], default_tipo: str):
    """Envía la foto de la libreta a Gemini y devuelve una lista estructurada de movimientos."""
    from google import genai

    api_key = get_gemini_api_key()
    if not api_key:
        raise ValueError("Falta configurar la API Key en gemini_config.json.")

    client = genai.Client(api_key=api_key)
    b64_img = base64.b64encode(image_bytes).decode("ascii")
    catalog_txt = "\n".join(f"- {name}" for name in catalog_names)

    prompt = (
        "Eres un asistente experto en inventario para el negocio 'Ají y Molino'. "
        "Lee cuidadosamente esta foto de una libreta escrita a mano con ventas/compras del día.\n"
        "Para cada línea que identifiques, devuelve un arreglo JSON válido (sin texto adicional) con objetos de esta forma:\n"
        '[{"texto_libreta": "lo que dice la libreta", "tipo": "SALIDA", "producto": "NOMBRE EXACTO DEL CATÁLOGO", "cantidad": 1, "nota": ""}]\n\n'
        "Reglas:\n"
        f"1. 'tipo' debe ser 'SALIDA' (venta), 'ENTRADA' (compra a proveedor) o 'AJUSTE'. Si no se especifica, usa '{default_tipo}'.\n"
        "2. 'producto' DEBE ser exactamente uno de los nombres de este catálogo (elige el más parecido):\n"
        f"{catalog_txt}\n"
        "3. Si no encuentras coincidencia clara en el catálogo, deja 'producto' como cadena vacía \"\".\n"
        "4. 'cantidad' debe ser un número entero positivo (por defecto 1 si no se ve cantidad).\n"
        "5. Responde ÚNICAMENTE con el JSON."
    )

    last_err = None
    raw_text = ""
    for model_name in ("gemini-3.8-flash", "gemini-3.5-flash-lite"):
        try:
            res = client.interactions.create(
                model=model_name,
                input=[
                    {"type": "text", "text": prompt},
                    {"type": "image", "data": b64_img, "mime_type": mime_type},
                ],
            )
            raw_text = (res.output_text or "").strip()
            break
        except Exception as e:
            last_err = e
            continue

    if not raw_text and last_err:
        raise last_err

    # Limpiar posibles bloques ```json ... ```
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_text, flags=re.IGNORECASE | re.DOTALL).strip()
    match = re.search(r"\[.*\]", cleaned, flags=re.DOTALL)
    if match:
        cleaned = match.group(0)
    data = json.loads(cleaned)
    if not isinstance(data, list):
        raise ValueError("La respuesta de Gemini no es una lista JSON válida.")
    return data


# ==============================================================================
# CONTROL DE ACCESO (EMPLEADO VS ADMINISTRADOR)
# ==============================================================================
def get_admin_pin():
    try:
        if "ADMIN_PIN" in st.secrets:
            return str(st.secrets["ADMIN_PIN"]).strip()
    except Exception:
        pass
    return os.environ.get("ADMIN_PIN", "1309").strip()


if "user_role" not in st.session_state:
    st.session_state["user_role"] = None
if "user_name" not in st.session_state:
    st.session_state["user_name"] = ""

if st.session_state["user_role"] is None:
    st.title("🌶️ Ají y Molino — Acceso al Sistema")
    st.caption("Selecciona tu perfil para continuar:")

    col_emp, col_adm = st.columns(2, gap="large")
    with col_emp:
        with st.container(border=True):
            st.subheader("👤 Modo Empleado / Vendedor")
            st.caption("Acceso rápido para registrar nuevas ventas y usar el carrito.")
            emp_name = st.text_input("Tu nombre", placeholder="Ej. Carlos, María...", key="login_emp_name")
            if st.button("Entrar como Vendedor", type="primary", width="stretch", key="btn_login_emp"):
                if not emp_name.strip():
                    st.warning("⚠️ Escribe tu nombre para registrar quién hace las ventas.")
                else:
                    st.session_state["user_role"] = "empleado"
                    st.session_state["user_name"] = emp_name.strip()
                    st.rerun()

    with col_adm:
        with st.container(border=True):
            st.subheader("🔐 Modo Administrador")
            st.caption("Acceso completo: Dashboard, Costos, Inventario, Libreta IA e Historial.")
            pin_in = st.text_input("PIN de Administrador", type="password", placeholder="Ingresa el PIN", key="login_admin_pin")
            if st.button("Entrar como Administrador", width="stretch", key="btn_login_adm"):
                if pin_in.strip() == get_admin_pin():
                    st.session_state["user_role"] = "admin"
                    st.session_state["user_name"] = "Administrador"
                    st.rerun()
                else:
                    st.error("❌ PIN incorrecto.")
    st.stop()

is_admin = st.session_state["user_role"] == "admin"

# ==============================================================================
# INTERFAZ GRÁFICA PRINCIPAL
# ==============================================================================
hdr_l, hdr_r = st.columns([4, 1.2])
with hdr_l:
    st.title("🌶️ Ají y Molino — Sistema de Inventario y Ventas")
    modo_conn = "☁️ Google Sheets (Nube 24/7)" if is_cloud_mode() else f"💻 Local (`{EXCEL_PATH.name}`)"
    rol_badge = "🔐 Administrador" if is_admin else f"👤 Vendedor: **{st.session_state['user_name']}** (Ventas e Inventario)"
    st.caption(f"{modo_conn}  ·  {rol_badge}")
with hdr_r:
    st.write("")
    if st.button("🚪 Cambiar Usuario", width="stretch"):
        st.session_state["user_role"] = None
        st.session_state["user_name"] = ""
        st.session_state["cart"] = []
        st.rerun()

df_inv, df_mov = load_data()


@st.dialog("✅ ¡Guardado con éxito!")
def mostrar_popup_guardado(mensaje: str):
    st.success(mensaje)
    if st.button("Aceptar", type="primary", width="stretch"):
        st.rerun()


if "popup_msg" in st.session_state:
    _msg = st.session_state.pop("popup_msg")
    st.toast(_msg, icon="✅")
    mostrar_popup_guardado(_msg)

if is_admin:
    tab_reg, tab_ocr, tab_dash, tab_inv, tab_hist = st.tabs(
        [
            "🛒 Registrar Movimiento",
            "📸 Leer Libreta (Foto)",
            "📊 Dashboard y Fechas",
            "📦 Inventario Completo",
            "📜 Historial de Movimientos",
        ]
    )
else:
    tab_reg, tab_inv = st.tabs(["🛒 Registrar Venta", "📦 Consultar Inventario"])
    tab_ocr = tab_dash = tab_hist = None

# ------------------------------------------------------------------------------
# TAB 1: REGISTRAR MOVIMIENTO (Admin: Todo | Empleado: Solo Ventas)
# ------------------------------------------------------------------------------
with tab_reg:
    st.subheader("Registrar Venta" if not is_admin else "Registrar Venta, Compra o Ajuste")

    if "cart" not in st.session_state:
        st.session_state["cart"] = []

    col_form, col_info = st.columns([1.3, 1], gap="large")

    with col_form:
        c1, c2 = st.columns(2)
        with c1:
            fecha_mov = st.date_input(
                "📅 Fecha",
                value=datetime.date.today(),
                format="DD/MM/YYYY",
                key="reg_fecha",
            )
        with c2:
            if is_admin:
                tipo_label = st.selectbox(
                    "🔄 Tipo de Movimiento",
                    ["SALIDA (Venta)", "ENTRADA (Compra a proveedor)", "AJUSTE (Conteo físico / Merma)"],
                    key="reg_tipo",
                )
                tipo_code = tipo_label.split(" ")[0]
            else:
                st.selectbox("🔄 Tipo de Movimiento", ["SALIDA (Venta)"], disabled=True, key="reg_tipo_emp")
                tipo_code = "SALIDA"

        productos_lista = sorted(df_inv["Producto"].tolist())
        producto_sel = st.selectbox(
            "🔍 Buscar Producto por Nombre (escribe para filtrar)",
            options=productos_lista,
            index=0,
            key="reg_prod",
        )

        prod_row = df_inv[df_inv["Producto"] == producto_sel].iloc[0]
        precio_alt = prod_row["Precio Alt."]
        has_alt = precio_alt is not None and not pd.isna(precio_alt) and int(precio_alt) > 0

        # Selector rápido de tarifa (Empleado nunca ve Costo c/IVA ni puede editar el precio)
        tarifa_opts = ["Precio Venta Normal (PVF)"]
        if has_alt:
            tarifa_opts.append(f"Precio Fracción / Mayor (${int(precio_alt):,})")
        if is_admin:
            tarifa_opts.append("Costo c/IVA")
        default_tarifa_idx = 0 if tipo_code == "SALIDA" else len(tarifa_opts) - 1

        c3, c4, c5 = st.columns([1, 1.3, 1])
        with c3:
            cantidad = st.number_input("🔢 Cantidad", min_value=1, step=1, value=1, key="reg_qty")
        with c4:
            tarifa_sel = st.selectbox("🏷️ Tarifa base", tarifa_opts, index=default_tarifa_idx, key="reg_tarifa")
        with c5:
            if "Fracción" in tarifa_sel and has_alt:
                precio_sugerido = int(precio_alt)
            elif "Costo" in tarifa_sel and is_admin:
                precio_sugerido = int(prod_row["Costo c/IVA"])
            else:
                precio_sugerido = int(prod_row["PVF"])
            precio_unitario = st.number_input(
                "💲 Valor Unitario ($)",
                min_value=0,
                step=100,
                value=precio_sugerido,
                disabled=not is_admin,
                key=f"reg_vunit_{producto_sel}_{tarifa_sel}",
            )

        c6, c7 = st.columns(2)
        with c6:
            pagos_disponibles = (
                [
                    "Efectivo",
                    "Nequi / Daviplata",
                    "Transferencia Bancaria",
                    "Tarjeta / Datáfono",
                    "Factura Proveedor",
                    "Crédito Cliente",
                    "Ajuste Interno",
                ]
                if is_admin
                else ["Efectivo", "Nequi / Daviplata", "Transferencia Bancaria", "Tarjeta / Datáfono"]
            )
            medio_pago = st.selectbox(
                "💳 Medio de Pago",
                pagos_disponibles,
                index=6 if (is_admin and tipo_code == "AJUSTE") else 0,
                key="reg_pago",
            )
        with c7:
            cliente_prov = st.text_input(
                "👤 Cliente" if not is_admin else "👤 Cliente / Proveedor / Motivo",
                placeholder="Ej. Venta mostrador, nombre cliente...",
                key="reg_cli",
            )

        ajuste_signo = "➕ Sumar"
        motivo_ajuste = ""
        if tipo_code == "AJUSTE":
            aj1, aj2 = st.columns(2)
            with aj1:
                ajuste_signo = st.radio(
                    "⚖️ Acción en el Stock",
                    ["➖ Restar (Dañado / Vencido / Merma)", "➕ Sumar (Sobrante)"],
                    horizontal=True,
                    key="reg_ajuste_signo",
                )
            with aj2:
                motivo_ajuste = st.selectbox(
                    "📌 Motivo del ajuste",
                    ["Producto dañado / echado a perder", "Producto vencido", "Merma / pérdida", "Conteo físico", "Otro"],
                    key="reg_ajuste_motivo",
                )

        obs_mov = st.text_input("📝 Observaciones (opcional)", key="reg_obs")

    cantidad_real = -int(cantidad) if (tipo_code == "AJUSTE" and "Restar" in ajuste_signo) else int(cantidad)
    total_mov = abs(int(cantidad_real)) * int(precio_unitario)
    stock_nuevo_est = (
        int(prod_row["Stock Actual"]) + cantidad_real
        if tipo_code in ("ENTRADA", "AJUSTE")
        else int(prod_row["Stock Actual"]) - int(cantidad)
    )

    with col_info:
        st.markdown("#### 📋 Detalle del Producto")
        m1, m2 = st.columns(2)
        m1.metric("Stock Actual", f"{prod_row['Stock Actual']} und", f"Quedará en: {stock_nuevo_est} und")
        m2.metric("SKU", prod_row["SKU"])

        m3, m4 = st.columns(2)
        m3.metric("Precio Venta (PVF)", f"${prod_row['PVF']:,}")
        if is_admin:
            m4.metric("Costo c/IVA", f"${prod_row['Costo c/IVA']:,}")
        else:
            m4.metric("Estado Stock", prod_row["Estado"])

        extra_alt = f"  \n**Precio Fracción / Mayor:** ${int(precio_alt):,}" if has_alt else ""
        st.info(
            f"**Categoría:** {prod_row['Categoría']}  \n"
            f"**Presentación:** {prod_row['Presentación']}{extra_alt}  \n"
            f"**Total de esta línea:** **${total_mov:,} COP**"
        )

        if (tipo_code == "SALIDA" and cantidad > prod_row["Stock Actual"]) or (
            tipo_code == "AJUSTE" and "Restar" in ajuste_signo and cantidad > prod_row["Stock Actual"]
        ):
            st.warning(
                f"⚠️ Atención: estás descontando **{cantidad}** und, pero solo hay **{prod_row['Stock Actual']}** en stock."
            )

    def _build_mov_notes():
        cli_out = cliente_prov.strip()
        if tipo_code == "AJUSTE" and not cli_out:
            cli_out = motivo_ajuste
        obs_out = obs_mov.strip()
        if not is_admin and st.session_state.get("user_name"):
            obs_out = f"[Vendedor: {st.session_state['user_name']}] {obs_out}".strip()
        if tipo_code == "AJUSTE" and motivo_ajuste and motivo_ajuste not in obs_out and cli_out != motivo_ajuste:
            obs_out = f"{motivo_ajuste} | {obs_out}".strip(" |")
        if precio_unitario != (prod_row["PVF"] if tipo_code == "SALIDA" else prod_row["Costo c/IVA"]):
            nota_precio = f"Tarifa aplicada: ${precio_unitario:,}"
            obs_out = f"{obs_out} | {nota_precio}".strip(" |")
        return cli_out, obs_out

    btn1, btn2 = st.columns(2)
    with btn1:
        if st.button("➕ Agregar al Carrito (Varios Productos)", width="stretch"):
            cli_out, obs_final = _build_mov_notes()
            st.session_state["cart"].append(
                {
                    "fecha": fecha_mov,
                    "tipo": tipo_code,
                    "producto": producto_sel,
                    "cantidad": int(cantidad_real),
                    "valor_unit": int(precio_unitario),
                    "total": int(total_mov),
                    "cliente": cli_out,
                    "pago": medio_pago,
                    "obs": obs_final,
                }
            )
            st.success(f"➕ Agregado al carrito: **{cantidad_real:+} x {producto_sel}** (${total_mov:,})")

    with btn2:
        if st.button("💾 Guardar Movimiento Directo", type="primary", width="stretch"):
            try:
                cli_out, obs_final = _build_mov_notes()
                fila_guardada = save_movement(
                    fecha_mov, tipo_code, producto_sel, cantidad_real, cli_out, medio_pago, obs_final
                )
                st.session_state.pop("hist_work", None)
                st.session_state["hist_ver"] = st.session_state.get("hist_ver", 0) + 1
                destino_txt = "Google Sheets" if is_cloud_mode() else "LISTADO_VENTA.xlsx"
                st.session_state["popup_msg"] = (
                    f"¡Movimiento guardado en {destino_txt} (fila {fila_guardada})!\n\n"
                    f"**{tipo_code}** · **{cantidad_real} x {producto_sel}** ({fecha_mov.strftime('%d/%m/%Y')}) — Stock nuevo: **{stock_nuevo_est} und**"
                )
                st.rerun()
            except PermissionError:
                st.error(
                    "❌ **No se pudo guardar porque tienes `LISTADO_VENTA.xlsx` abierto en Excel.** "
                    "Ciérralo en Excel un momento y vuelve a hacer clic en Guardar."
                )
            except Exception as e:
                st.error(f"❌ Error al guardar: {e}")

    # Mostrar Carrito si tiene ítems
    if st.session_state["cart"]:
        st.divider()
        st.markdown(f"### 🛍️ Carrito Actual ({len(st.session_state['cart'])} productos)")
        df_cart = pd.DataFrame(st.session_state["cart"])
        total_carrito = int(df_cart["total"].sum())
        st.dataframe(df_cart, width="stretch", hide_index=True)
        st.markdown(f"#### 💰 Total del Pedido: **${total_carrito:,} COP**")

        # Recibo rápido para WhatsApp
        lineas_wsp = [
            f"🌶️ *AJÍ Y MOLINO* — Resumen ({fecha_mov.strftime('%d/%m/%Y')})",
            *(f"• {it['cantidad']}x {it['producto']} — ${it['total']:,}" for it in st.session_state["cart"]),
            f"*TOTAL: ${total_carrito:,} COP* ({medio_pago})",
        ]
        with st.expander("📲 Copiar recibo para WhatsApp"):
            st.code("\n".join(lineas_wsp), language="text")

        cb1, cb2 = st.columns([3, 1])
        with cb1:
            if st.button("✅ Guardar Todo el Carrito en Excel", type="primary", width="stretch"):
                try:
                    filas = save_cart_movements(st.session_state["cart"])
                    st.session_state["cart"] = []
                    st.session_state.pop("hist_work", None)
                    st.session_state["hist_ver"] = st.session_state.get("hist_ver", 0) + 1
                    destino_txt = "Google Sheets" if is_cloud_mode() else "LISTADO_VENTA.xlsx"
                    st.session_state["popup_msg"] = f"¡Carrito guardado en {destino_txt} (filas {filas[0]} a {filas[-1]})!"
                    st.rerun()
                except PermissionError:
                    st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder guardar el carrito.")
                except Exception as e:
                    st.error(f"❌ Error al guardar el carrito: {e}")
        with cb2:
            if st.button("🗑️ Vaciar Carrito", width="stretch"):
                st.session_state["cart"] = []
                st.rerun()

# ------------------------------------------------------------------------------
# TAB 3: INVENTARIO COMPLETO + EDICIÓN RÁPIDA DE PRECIOS/STOCK
# ------------------------------------------------------------------------------
with tab_inv:
    st.subheader("Catálogo e Inventario Actual (138+ productos)")

    # Bloque optimizado para celular: botones grandes sin tener que editar tablas anchas
    if is_admin:
        with st.container(border=True):
            st.markdown("#### 📱 Cambio Rápido de Stock (Modo Celular)")
            st.caption("Ideal desde el celular: busca el producto, ajusta el número con **- / +** y guarda en un toque.")
            m_prod = st.selectbox(
                "🔍 Producto",
                sorted(df_inv["Producto"].tolist()),
                key="mob_stock_prod",
            )
            m_row = df_inv[df_inv["Producto"] == m_prod].iloc[0]

            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("Stock Actual", f"{int(m_row['Stock Actual'])} und", m_row["Estado"])
            mc2.metric("Stock Inicial", f"{int(m_row['Stock Inicial'])} und")
            mc3.metric("Stock Mínimo", f"{int(m_row['Stock Mín.'])} und")

            m_modo = st.radio(
                "¿Qué quieres ajustar?",
                ["🎯 Stock Actual (Conteo físico)", "📦 Stock Inicial directamente"],
                horizontal=True,
                key="mob_stock_modo",
            )

            val_def = int(m_row["Stock Actual"]) if "Actual" in m_modo else int(m_row["Stock Inicial"])
            nuevo_val = st.number_input(
                "Nueva cantidad",
                min_value=0,
                step=1,
                value=val_def,
                key=f"mob_val_{m_prod}_{m_modo}",
            )

            if st.button("✅ Guardar Nuevo Stock", type="primary", width="stretch", key="mob_btn_save"):
                try:
                    if "Actual" in m_modo:
                        # Calcular el Stock Inicial necesario para que Stock Actual quede exactamente en nuevo_val
                        # (o si prefieres ajuste, aquí mantiene Entradas/Salidas intactas ajustando Stock Inicial)
                        diff = int(nuevo_val) - int(m_row["Stock Actual"])
                        new_s_ini = max(0, int(m_row["Stock Inicial"]) + diff)
                        if int(m_row["Stock Inicial"]) + diff < 0:
                            # Si quedaría negativo el inicial, registramos un AJUSTE en MOVIMIENTOS
                            save_movement(
                                datetime.date.today(),
                                "AJUSTE",
                                m_prod,
                                diff,
                                "Conteo físico celular",
                                "Ajuste Interno",
                                f"Ajuste de {int(m_row['Stock Actual'])} a {int(nuevo_val)}",
                            )
                        else:
                            update_inventory_bulk([
                                {
                                    "fila_excel": m_row["fila_excel"],
                                    "stock_ini": new_s_ini,
                                    "stock_min": int(m_row["Stock Mín."]),
                                    "costo_base": int(m_row["Costo Base"]),
                                    "pvf": int(m_row["PVF"]),
                                    "obs": str(m_row["Observaciones"] or ""),
                                }
                            ])
                        st.session_state["popup_msg"] = f"Stock Actual de **{m_prod}** actualizado a **{int(nuevo_val)} und**."
                    else:
                        update_inventory_bulk([
                            {
                                "fila_excel": m_row["fila_excel"],
                                "stock_ini": int(nuevo_val),
                                "stock_min": int(m_row["Stock Mín."]),
                                "costo_base": int(m_row["Costo Base"]),
                                "pvf": int(m_row["PVF"]),
                                "obs": str(m_row["Observaciones"] or ""),
                            }
                        ])
                        st.session_state["popup_msg"] = f"Stock Inicial de **{m_prod}** actualizado a **{int(nuevo_val)} und**."
                    st.session_state.pop("hist_work", None)
                    st.session_state["hist_ver"] = st.session_state.get("hist_ver", 0) + 1
                    st.rerun()
                except PermissionError:
                    st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder guardar el cambio.")
                except Exception as e:
                    st.error(f"❌ Error al actualizar stock: {e}")

    st.divider()
    ci1, ci2, ci3 = st.columns([2, 1.5, 1])
    with ci1:
        buscar_txt = st.text_input("🔍 Buscar por nombre o SKU", placeholder="Ej. habanero, mango, BO01001...")
    with ci2:
        cat_sel = st.selectbox("Categoría", ["(Todas)"] + sorted(df_inv["Categoría"].unique().tolist()))
    with ci3:
        est_sel = st.selectbox("Estado", ["(Todos)", "🟢 OK", "🟡 PEDIR", "🔴 AGOTADO"])

    df_view = df_inv.copy()
    if buscar_txt.strip():
        q = buscar_txt.strip().lower()
        df_view = df_view[
            df_view["Producto"].str.lower().str.contains(q, na=False)
            | df_view["SKU"].str.lower().str.contains(q, na=False)
        ]
    if cat_sel != "(Todas)":
        df_view = df_view[df_view["Categoría"] == cat_sel]
    if est_sel != "(Todos)":
        df_view = df_view[df_view["Estado"] == est_sel]

    if not is_admin:
        cols_emp = ["SKU", "Categoría", "Producto", "Presentación", "Stock Actual", "Estado", "PVF", "Observaciones"]
        st.dataframe(df_view[cols_emp], width="stretch", hide_index=True, height=480)
    else:
        cols_mostrar = [
            "SKU", "Categoría", "Producto", "Presentación",
            "Stock Inicial", "Entradas (+)", "Salidas (-)", "Ajustes (+/-)",
            "Stock Actual", "Stock Mín.", "Estado", "Costo Base", "Costo c/IVA", "PVF",
            "Valor Inv. Venta", "Observaciones",
        ]
        st.caption("✏️ Puedes modificar directamente en la tabla **Stock Inicial**, **Stock Mín.**, **Costo Base**, **PVF** u **Observaciones** y luego guardar.")
        df_inv_edit_src = df_view[cols_mostrar].copy()
        edited_inv_df = st.data_editor(
            df_inv_edit_src,
            width="stretch",
            hide_index=True,
            height=460,
            disabled=[
                "SKU", "Categoría", "Producto", "Presentación",
                "Entradas (+)", "Salidas (-)", "Ajustes (+/-)",
                "Stock Actual", "Estado", "Costo c/IVA", "Valor Inv. Venta",
            ],
            column_config={
                "Stock Inicial": st.column_config.NumberColumn("Stock Inicial ✏️", min_value=0, step=1),
                "Stock Mín.": st.column_config.NumberColumn("Stock Mín. ✏️", min_value=0, step=1),
                "Costo Base": st.column_config.NumberColumn("Costo Base ✏️", min_value=0, step=100, format="$%d"),
                "Costo c/IVA": st.column_config.NumberColumn("Costo c/IVA", format="$%d"),
                "PVF": st.column_config.NumberColumn("PVF ✏️", min_value=0, step=500, format="$%d"),
                "Valor Inv. Venta": st.column_config.NumberColumn("Valor Inv. Venta", format="$%d"),
                "Observaciones": st.column_config.TextColumn("Observaciones ✏️"),
            },
            key="inv_bulk_editor",
        )

        if st.button("💾 Guardar Cambios de Inventario Inicial / Precios", type="primary", width="stretch"):
            changes = []
            orig_by_prod = df_view.set_index("Producto").to_dict(orient="index")
            for _, r_ed in edited_inv_df.iterrows():
                prod_k = str(r_ed["Producto"])
                orig = orig_by_prod.get(prod_k)
                if not orig:
                    continue
                s_ini = int(round(to_float(r_ed["Stock Inicial"], 0)))
                s_min = int(round(to_float(r_ed["Stock Mín."], 0)))
                c_bas = int(round(to_float(r_ed["Costo Base"], 0)))
                pvf_v = int(round(to_float(r_ed["PVF"], 0)))
                obs_v = "" if pd.isna(r_ed["Observaciones"]) else str(r_ed["Observaciones"] or "").strip()

                if (
                    s_ini != int(orig["Stock Inicial"])
                    or s_min != int(orig["Stock Mín."])
                    or c_bas != int(orig["Costo Base"])
                    or pvf_v != int(orig["PVF"])
                    or obs_v != str(orig["Observaciones"] or "").strip()
                ):
                    changes.append(
                        {
                            "fila_excel": orig["fila_excel"],
                            "stock_ini": s_ini,
                            "stock_min": s_min,
                            "costo_base": c_bas,
                            "pvf": pvf_v,
                            "obs": obs_v,
                        }
                    )
            if not changes:
                st.info("No hay cambios pendientes por guardar en la tabla.")
            else:
                try:
                    n_upd = update_inventory_bulk(changes)
                    st.session_state["popup_msg"] = f"¡Se actualizaron **{n_upd} producto(s)** en el inventario!"
                    st.rerun()
                except PermissionError:
                    st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder guardar los cambios.")
                except Exception as e:
                    st.error(f"❌ Error al guardar cambios: {e}")

        with st.expander("✏️ Editar Costo, Precio de Venta (PVF) o Stock Mínimo de un producto"):
            prod_edit = st.selectbox("Selecciona el producto a editar", sorted(df_inv["Producto"].tolist()), key="edit_prod")
            row_e = df_inv[df_inv["Producto"] == prod_edit].iloc[0]
            e1, e2, e3, e4 = st.columns(4)
            new_sini = e1.number_input("Stock Inicial", value=int(row_e["Stock Inicial"]), step=1)
            new_smin = e2.number_input("Stock Mínimo", value=int(row_e["Stock Mín."]), step=1)
            new_cbase = e3.number_input("Costo Base (sin IVA)", value=int(row_e["Costo Base"]), step=100)
            new_pvf = e4.number_input("Precio Venta Final (PVF)", value=int(row_e["PVF"]), step=500)
            new_obs = st.text_input("Observaciones", value=str(row_e["Observaciones"]))

            if st.button("💾 Guardar Cambios del Producto"):
                try:
                    update_product_in_excel(row_e["fila_excel"], new_sini, new_smin, new_cbase, new_pvf, new_obs)
                    st.session_state["popup_msg"] = f"Producto **{prod_edit}** actualizado correctamente."
                    st.rerun()
                except PermissionError:
                    st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder guardar los cambios.")

        with st.expander("➕ Crear Nuevo Producto en el Catálogo"):
            n1, n2, n3 = st.columns([1, 1.5, 1])
            new_sku = n1.text_input("SKU / Código *", placeholder="Ej. AJI-003", key="new_sku")
            cats_existentes = sorted(df_inv["Categoría"].unique().tolist())
            new_cat = n2.selectbox("Categoría *", cats_existentes, key="new_cat")
            new_pres = n3.text_input("Presentación", value="und", key="new_pres")

            new_desc = st.text_input("Nombre / Descripción del Producto *", placeholder="Ej. SALSA HABANERO MANGO 250ML", key="new_desc")

            n4, n5, n6, n7, n8 = st.columns(5)
            new_p_sini = n4.number_input("Stock Inicial", min_value=0, step=1, value=0, key="new_p_sini")
            new_p_smin = n5.number_input("Stock Mínimo", min_value=0, step=1, value=3, key="new_p_smin")
            new_p_cbase = n6.number_input("Costo Base (sin IVA)", min_value=0, step=100, value=0, key="new_p_cbase")
            new_p_iva = n7.selectbox("IVA %", [0.19, 0.0, 0.05], format_func=lambda x: f"{int(x*100)}%", key="new_p_iva")
            new_p_pvf = n8.number_input("Precio Venta Final (PVF)", min_value=0, step=500, value=0, key="new_p_pvf")
            new_p_obs = st.text_input("Observaciones", key="new_p_obs")

            if st.button("➕ Agregar Producto al Inventario", type="primary"):
                if not new_sku.strip() or not new_desc.strip():
                    st.warning("⚠️ El SKU y el Nombre del Producto son obligatorios.")
                elif new_desc.strip().upper() in df_inv["Producto"].str.upper().values:
                    st.warning("⚠️ Ya existe un producto con ese mismo nombre.")
                elif new_sku.strip().upper() in df_inv["SKU"].str.upper().values:
                    st.warning("⚠️ Ya existe un producto con ese mismo SKU.")
                else:
                    try:
                        fila_nueva = create_new_product_in_excel(
                            new_sku, new_cat, new_desc, new_pres,
                            new_p_sini, new_p_smin, new_p_cbase, new_p_iva, 0.30, new_p_pvf, new_p_obs
                        )
                        st.session_state["popup_msg"] = f"Producto **{new_desc.strip()}** agregado al catálogo (fila {fila_nueva})."
                        st.rerun()
                    except PermissionError:
                        st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder agregar el producto.")
                    except Exception as e:
                        st.error(f"❌ Error al crear el producto: {e}")

if not is_admin:
    st.stop()

# ------------------------------------------------------------------------------
# TAB 2: DASHBOARD Y CONSULTA POR RANGO DE FECHAS
# ------------------------------------------------------------------------------
with tab_dash:
    st.subheader("Resumen General del Negocio")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Productos (SKUs)", f"{len(df_inv):,}")
    k2.metric("Unidades en Stock", f"{int(df_inv['Stock Actual'].sum()):,}")
    k3.metric("Valor Inv. al Costo", f"${int(df_inv['Valor Inv. Costo'].sum()):,}")
    k4.metric("Valor Inv. a Venta", f"${int(df_inv['Valor Inv. Venta'].sum()):,}")

    k5, k6, k7, k8 = st.columns(4)
    util_pot = int(df_inv["Valor Inv. Venta"].sum() - df_inv["Valor Inv. Costo"].sum())
    k5.metric("Utilidad Bruta Potencial", f"${util_pot:,}")
    k6.metric("🟢 Productos OK", int((df_inv["Estado"] == "🟢 OK").sum()))
    k7.metric("🟡 Por Pedir", int((df_inv["Estado"] == "🟡 PEDIR").sum()))
    k8.metric("🔴 Agotados", int((df_inv["Estado"] == "🔴 AGOTADO").sum()))

    st.divider()
    st.subheader("📅 Consulta de Compras y Ventas por Rango de Fechas")

    f1, f2, f3, f4 = st.columns([1, 1, 1.5, 1])
    with f1:
        f_desde = st.date_input("Desde", value=datetime.date.today().replace(day=1), format="DD/MM/YYYY", key="dash_desde")
    with f2:
        f_hasta = st.date_input("Hasta", value=datetime.date.today(), format="DD/MM/YYYY", key="dash_hasta")
    with f3:
        prod_filtro = st.selectbox("Producto (Opcional)", ["(Todos los productos)"] + sorted(df_inv["Producto"].tolist()), key="dash_prod")
    with f4:
        tipo_filtro = st.selectbox("Tipo (Opcional)", ["(Todos)", "ENTRADA", "SALIDA", "AJUSTE"], key="dash_tipo")

    if not df_mov.empty:
        df_rango = df_mov.copy()
        df_rango = df_rango[df_rango["Fecha"].notna()]
        df_rango = df_rango[(df_rango["Fecha"] >= f_desde) & (df_rango["Fecha"] <= f_hasta)]
        if prod_filtro != "(Todos los productos)":
            df_rango = df_rango[df_rango["Producto"] == prod_filtro]
        if tipo_filtro != "(Todos)":
            df_rango = df_rango[df_rango["Tipo"] == tipo_filtro]

        c_ent = df_rango[df_rango["Tipo"] == "ENTRADA"]
        c_sal = df_rango[df_rango["Tipo"] == "SALIDA"]
        c_aju = df_rango[df_rango["Tipo"] == "AJUSTE"]

        r1, r2, r3, r4, r5 = st.columns(5)
        r1.metric("Compras (Und)", f"{int(c_ent['Cantidad'].sum()):,}")
        r2.metric("Compras ($ COP)", f"${int(c_ent['Total'].sum()):,}")
        r3.metric("Ventas (Und)", f"{int(c_sal['Cantidad'].sum()):,}")
        r4.metric("Ventas ($ COP)", f"${int(c_sal['Total'].sum()):,}")
        r5.metric("Ajustes / Dañados (Und)", f"{int(c_aju['Cantidad'].sum()):,}")

        if not df_rango.empty:
            st.dataframe(
                df_rango.drop(columns=["fila_excel"]),
                width="stretch",
                hide_index=True,
            )

            # Cuadre de Caja por Medio de Pago en el rango seleccionado
            if not c_sal.empty:
                st.markdown("##### 💳 Cuadre de Ventas por Medio de Pago (en el rango)")
                df_cuadre = (
                    c_sal.assign(**{"Medio de Pago": c_sal["Medio de Pago"].replace("", "Sin especificar")})
                    .groupby("Medio de Pago", as_index=False)
                    .agg(Operaciones=("Cantidad", "count"), Unidades=("Cantidad", "sum"), Total_COP=("Total", "sum"))
                )
                st.dataframe(df_cuadre, width="stretch", hide_index=True)
        else:
            st.info("No hay movimientos registrados con fecha dentro de ese rango.")
    else:
        st.info("Aún no hay movimientos registrados.")

    st.divider()
    st.subheader("🚨 Productos Agotados o Por Pedir (Lista de Reposición)")
    df_alertas = df_inv[df_inv["Estado"] != "🟢 OK"][
        ["SKU", "Categoría", "Producto", "Stock Actual", "Stock Mín.", "Estado", "Costo c/IVA"]
    ].copy()
    if df_alertas.empty:
        st.success("🎉 ¡Todo el inventario está en verde (🟢 OK)!")
    else:
        st.dataframe(df_alertas, width="stretch", hide_index=True)
        csv_alertas = df_alertas.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📥 Descargar Lista de Pedido (CSV)",
            data=csv_alertas,
            file_name=f"Lista_Pedido_{datetime.date.today().strftime('%Y-%m-%d')}.csv",
            mime="text/csv",
        )

    st.divider()
    st.subheader("📦 Valorización por Categoría")
    df_cat = (
        df_inv.groupby("Categoría", as_index=False)
        .agg(
            SKUs=("SKU", "count"),
            Unidades=("Stock Actual", "sum"),
            Costo_Total=("Valor Inv. Costo", "sum"),
            Venta_Total=("Valor Inv. Venta", "sum"),
        )
    )
    df_cat["Utilidad_Potencial"] = df_cat["Venta_Total"] - df_cat["Costo_Total"]
    st.dataframe(df_cat, width="stretch", hide_index=True)
    st.bar_chart(df_cat.set_index("Categoría")[["Costo_Total", "Venta_Total"]])

# ------------------------------------------------------------------------------
# TAB 4: HISTORIAL INTERACTIVO (TIPO EXCEL — CON FRAGMENT ULTRA-RÁPIDO)
# ------------------------------------------------------------------------------
_fragment_decorator = getattr(st, "fragment", lambda f: f)


@_fragment_decorator
def render_historial_interactivo(df_inv_local, df_mov_local):
    if "hist_ver" not in st.session_state:
        st.session_state["hist_ver"] = 0

    if "hist_work" not in st.session_state:
        if df_mov_local.empty:
            st.session_state["hist_work"] = pd.DataFrame(
                columns=[
                    "ID Lote", "Fecha", "Tipo", "Producto", "SKU", "Categoría",
                    "Cantidad", "Valor Unitario", "Total", "Cliente / Proveedor",
                    "Medio de Pago", "Observaciones",
                ]
            )
        else:
            st.session_state["hist_work"] = df_mov_local.drop(columns=["fila_excel"]).copy()

    df_edit_base = st.session_state["hist_work"]

    productos_opts = sorted(df_inv_local["Producto"].tolist())
    pagos_opts = [
        "",
        "Efectivo",
        "Nequi / Daviplata",
        "Transferencia Bancaria",
        "Tarjeta / Datáfono",
        "Factura Proveedor",
        "Crédito Cliente",
        "Ajuste Interno",
    ]

    edited_df = st.data_editor(
        df_edit_base,
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        height=440,
        disabled=["ID Lote", "SKU", "Categoría", "Valor Unitario", "Total"],
        column_config={
            "ID Lote": st.column_config.TextColumn("ID Lote (Auto)", width="small"),
            "Fecha": st.column_config.DateColumn("📅 Fecha", format="DD/MM/YYYY", width="small"),
            "Tipo": st.column_config.SelectboxColumn(
                "Tipo", options=["SALIDA", "ENTRADA", "AJUSTE"], width="small"
            ),
            "Producto": st.column_config.SelectboxColumn(
                "Producto", options=productos_opts, width="large"
            ),
            "SKU": st.column_config.TextColumn("SKU (Auto)", width="small"),
            "Categoría": st.column_config.TextColumn("Categoría (Auto)", width="medium"),
            "Cantidad": st.column_config.NumberColumn("Cantidad", step=1, width="small"),
            "Valor Unitario": st.column_config.NumberColumn("Valor Unit. (Auto)", format="$%d", width="small"),
            "Total": st.column_config.NumberColumn("Total (Auto)", format="$%d", width="small"),
            "Cliente / Proveedor": st.column_config.TextColumn("Cliente / Proveedor", width="medium"),
            "Medio de Pago": st.column_config.SelectboxColumn("Medio de Pago", options=pagos_opts, width="medium"),
            "Observaciones": st.column_config.TextColumn("Observaciones", width="medium"),
        },
        key=f"hist_editor_{st.session_state['hist_ver']}",
    )

    prod_lookup = df_inv_local.set_index("Producto").to_dict(orient="index")
    needs_refresh = False
    recalc_df = edited_df.copy()

    for idx_r, row_r in recalc_df.iterrows():
        prod_name = "" if pd.isna(row_r.get("Producto")) else str(row_r.get("Producto") or "").strip()
        if prod_name.lower() in ("nan", "none"):
            prod_name = ""
        tipo_val = "" if pd.isna(row_r.get("Tipo")) else str(row_r.get("Tipo") or "").strip().upper()
        if tipo_val in ("NAN", "NONE", "⚠️ SIN TIPO"):
            tipo_val = ""
        fecha_val = to_date(row_r.get("Fecha"))
        qty_val = int(round(to_float(row_r.get("Cantidad"), 0)))

        if prod_name and prod_name in prod_lookup:
            if not tipo_val:
                tipo_val = "SALIDA"
                recalc_df.at[idx_r, "Tipo"] = "SALIDA"
                needs_refresh = True
            if fecha_val is None:
                fecha_val = datetime.date.today()
                recalc_df.at[idx_r, "Fecha"] = fecha_val
                needs_refresh = True
            if qty_val == 0:
                qty_val = 1
                recalc_df.at[idx_r, "Cantidad"] = 1
                needs_refresh = True

            info_p = prod_lookup[prod_name]
            exp_sku = str(info_p["SKU"])
            exp_cat = str(info_p["Categoría"])
            exp_vunit = int(info_p["PVF"] if tipo_val == "SALIDA" else info_p["Costo c/IVA"])
            exp_tot = int(qty_val * exp_vunit)
            exp_lote = f"MOV-{idx_r + 1:03d}" + (f"-{fecha_val.strftime('%d%m')}" if fecha_val else "")
        else:
            exp_sku, exp_cat, exp_vunit, exp_tot, exp_lote = "", "", 0, 0, ""

        cur_sku = "" if pd.isna(row_r.get("SKU")) else str(row_r.get("SKU") or "")
        cur_cat = "" if pd.isna(row_r.get("Categoría")) else str(row_r.get("Categoría") or "")
        cur_vunit = int(round(to_float(row_r.get("Valor Unitario"), 0)))
        cur_tot = int(round(to_float(row_r.get("Total"), 0)))
        cur_lote = "" if pd.isna(row_r.get("ID Lote")) else str(row_r.get("ID Lote") or "")

        if (cur_sku != exp_sku) or (cur_cat != exp_cat) or (cur_vunit != exp_vunit) or (cur_tot != exp_tot) or (cur_lote != exp_lote):
            recalc_df.at[idx_r, "SKU"] = exp_sku
            recalc_df.at[idx_r, "Categoría"] = exp_cat
            recalc_df.at[idx_r, "Valor Unitario"] = exp_vunit
            recalc_df.at[idx_r, "Total"] = exp_tot
            recalc_df.at[idx_r, "ID Lote"] = exp_lote
            needs_refresh = True

    if needs_refresh:
        st.session_state["hist_work"] = recalc_df
        st.session_state["hist_ver"] += 1
        st.rerun(scope="fragment" if hasattr(st, "fragment") else "app")

    col_b1, col_b2, col_b3 = st.columns([2.2, 1.3, 1.3])
    with col_b1:
        if st.button("💾 Guardar Cambios de la Tabla", type="primary", width="stretch"):
            try:
                save_movements_table(edited_df)
                st.session_state.pop("hist_work", None)
                st.session_state["hist_ver"] += 1
                st.session_state["popup_msg"] = "¡Tabla de movimientos guardada! Stock y totales recalculados."
                st.rerun()
            except PermissionError:
                st.error(
                    "❌ **No se pudo guardar porque tienes `LISTADO_VENTA.xlsx` abierto en Excel.** "
                    "Ciérralo en Excel y vuelve a presionar Guardar."
                )
            except Exception as e:
                st.error(f"❌ Error al guardar cambios: {e}")
    with col_b2:
        if st.button("🔄 Recargar", width="stretch"):
            _clear_all_caches()
            st.session_state.pop("hist_work", None)
            st.session_state["hist_ver"] += 1
            st.rerun()
    with col_b3:
        if st.button(
            "🗑️ Descartar Movimientos...",
            width="stretch",
            disabled=df_mov_local.empty,
        ):
            mostrar_popup_descartar_movimientos(df_mov_local)


@st.dialog("🗑️ Seleccionar movimientos a descartar")
def mostrar_popup_descartar_movimientos(df_mov_local):
    st.caption("Marca la casilla de los movimientos que quieres eliminar y presiona **Confirmar Eliminación**.")
    sel_all = st.checkbox("Seleccionar todos", key="pop_del_all")

    indices_a_borrar = []
    for i, r in df_mov_local.iterrows():
        lbl = f"**{r['ID Lote']}** · {r['Fecha'] or 'Sin fecha'} · `{r['Tipo']}` · **{r['Cantidad']} x {r['Producto']}** (${int(r['Total']):,})"
        if st.checkbox(lbl, value=sel_all, key=f"pop_chk_{i}_{sel_all}"):
            indices_a_borrar.append(i)

    st.divider()
    b_conf, b_canc = st.columns(2)
    with b_conf:
        if st.button(
            f"🗑️ Eliminar ({len(indices_a_borrar)})",
            type="primary",
            width="stretch",
            disabled=(len(indices_a_borrar) == 0),
        ):
            try:
                df_restante = df_mov_local.drop(index=indices_a_borrar).drop(columns=["fila_excel"])
                save_movements_table(df_restante)
                st.session_state.pop("hist_work", None)
                st.session_state["hist_ver"] = st.session_state.get("hist_ver", 0) + 1
                st.session_state["popup_msg"] = f"Se eliminaron **{len(indices_a_borrar)} movimiento(s)** correctamente."
                st.rerun()
            except PermissionError:
                st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder eliminar.")
            except Exception as e:
                st.error(f"❌ Error al eliminar: {e}")
    with b_canc:
        if st.button("Cancelar", width="stretch"):
            st.rerun()


with tab_hist:
    st.subheader("Historial Interactivo de Entradas, Salidas y Ajustes")
    st.caption(
        "Haz doble clic en cualquier celda para editar **Fecha (calendario)**, **Tipo**, **Producto**, **Cantidad**, "
        "**Cliente/Proveedor**, **Medio de Pago** u **Observaciones**. También puedes **agregar filas abajo** o **descartar movimientos con casillas**."
    )
    render_historial_interactivo(df_inv, df_mov)


# ------------------------------------------------------------------------------
# TAB 5: 📸 LEER LIBRETA CON GEMINI (FOTO -> EXCEL)
# ------------------------------------------------------------------------------
with tab_ocr:
    st.subheader("📸 Pasar Libreta de Tu Papá a Excel Automáticamente")
    st.caption(
        "Sube la foto de la libreta del día. La IA lee la letra a mano, la empareja con tus 138 productos "
        "y te muestra una tabla para confirmar antes de guardar."
    )

    o1, o2, o3 = st.columns(3)
    with o1:
        ocr_fecha = st.date_input("📅 Fecha de la libreta", value=datetime.date.today(), format="DD/MM/YYYY", key="ocr_fecha")
    with o2:
        ocr_tipo_def = st.selectbox("🔄 Tipo por defecto", ["SALIDA", "ENTRADA"], key="ocr_tipo_def")
    with o3:
        ocr_pago_def = st.selectbox(
            "💳 Medio de pago por defecto",
            ["Efectivo", "Nequi / Daviplata", "Transferencia Bancaria", "Tarjeta / Datáfono", "Factura Proveedor"],
            key="ocr_pago_def",
        )

    modo_img = st.radio(
        "¿Cómo quieres cargar la foto?",
        ["📤 Subir foto de la galería", "📷 Tomar foto con la cámara del celular/PC"],
        horizontal=True,
        key="ocr_modo_img",
    )
    if "Subir" in modo_img:
        img_file = st.file_uploader("📤 Sube la foto de la libreta (JPG, PNG, WEBP)", type=["jpg", "jpeg", "png", "webp"], key="ocr_file")
    else:
        img_file = st.camera_input("📷 Toma la foto de la libreta", key="ocr_cam")

    if img_file is not None:
        col_img, col_act = st.columns([1, 1.5], gap="large")
        with col_img:
            st.image(img_file, caption="Foto cargada de la libreta", width="stretch")
        with col_act:
            if st.button("🤖 Leer Libreta con IA (Gemini)", type="primary", width="stretch"):
                with st.spinner("Leyendo letra a mano y buscando productos en tu catálogo..."):
                    try:
                        img_bytes = img_file.getvalue()
                        mime_t = img_file.type or "image/jpeg"
                        parsed = parse_notebook_image_with_gemini(
                            img_bytes, mime_t, sorted(df_inv["Producto"].tolist()), ocr_tipo_def
                        )
                        prod_map_ocr = df_inv.set_index("Producto").to_dict(orient="index")
                        rows_ocr = []
                        for item in parsed:
                            p_name = str(item.get("producto") or "").strip()
                            t_code = str(item.get("tipo") or ocr_tipo_def).strip().upper()
                            if t_code not in ("SALIDA", "ENTRADA", "AJUSTE"):
                                t_code = ocr_tipo_def
                            qty_i = max(1, int(round(to_float(item.get("cantidad"), 1))))
                            info_o = prod_map_ocr.get(p_name, {})
                            v_u = int(info_o.get("PVF", 0) if t_code == "SALIDA" else info_o.get("Costo c/IVA", 0))
                            rows_ocr.append(
                                {
                                    "Texto Libreta": str(item.get("texto_libreta") or ""),
                                    "Tipo": t_code,
                                    "Producto": p_name if p_name in prod_map_ocr else None,
                                    "Cantidad": qty_i,
                                    "Valor Unit.": v_u,
                                    "Total": qty_i * v_u,
                                    "Medio de Pago": ocr_pago_def,
                                    "Nota / Cliente": str(item.get("nota") or ""),
                                }
                            )
                        st.session_state["ocr_rows"] = pd.DataFrame(rows_ocr)
                        st.session_state["ocr_ver"] = st.session_state.get("ocr_ver", 0) + 1
                        st.success(f"✅ Se detectaron **{len(rows_ocr)}** líneas en la libreta. Revísalas abajo antes de guardar.")
                    except Exception as e:
                        st.error(f"❌ No se pudo leer la foto: {e}")

    if "ocr_rows" in st.session_state and not st.session_state["ocr_rows"].empty:
        st.divider()
        st.markdown("### ✏️ Revisa y confirma antes de guardar en Excel")
        productos_opts_ocr = sorted(df_inv["Producto"].tolist())
        pagos_opts_ocr = [
            "Efectivo", "Nequi / Daviplata", "Transferencia Bancaria",
            "Tarjeta / Datáfono", "Factura Proveedor", "Crédito Cliente", "Ajuste Interno",
        ]

        edited_ocr = st.data_editor(
            st.session_state["ocr_rows"],
            width="stretch",
            hide_index=True,
            num_rows="dynamic",
            disabled=["Texto Libreta", "Valor Unit.", "Total"],
            column_config={
                "Texto Libreta": st.column_config.TextColumn("Lo que decía la libreta", width="medium"),
                "Tipo": st.column_config.SelectboxColumn("Tipo", options=["SALIDA", "ENTRADA", "AJUSTE"], width="small"),
                "Producto": st.column_config.SelectboxColumn("Producto en Catálogo", options=productos_opts_ocr, width="large"),
                "Cantidad": st.column_config.NumberColumn("Cantidad", step=1, width="small"),
                "Valor Unit.": st.column_config.NumberColumn("Valor Unit.", format="$%d", width="small"),
                "Total": st.column_config.NumberColumn("Total", format="$%d", width="small"),
                "Medio de Pago": st.column_config.SelectboxColumn("Medio de Pago", options=pagos_opts_ocr, width="medium"),
                "Nota / Cliente": st.column_config.TextColumn("Nota / Cliente", width="medium"),
            },
            key=f"ocr_editor_{st.session_state.get('ocr_ver', 0)}",
        )

        # Recalcular en vivo si el usuario cambia Producto, Tipo o Cantidad en la tabla de revisión
        prod_map_ocr = df_inv.set_index("Producto").to_dict(orient="index")
        need_ocr_refresh = False
        recalc_ocr = edited_ocr.copy()
        for ir, rr in recalc_ocr.iterrows():
            pn = "" if pd.isna(rr.get("Producto")) else str(rr.get("Producto") or "").strip()
            tv = "" if pd.isna(rr.get("Tipo")) else str(rr.get("Tipo") or ocr_tipo_def).strip().upper()
            qv = max(1, int(round(to_float(rr.get("Cantidad"), 1))))
            info_o = prod_map_ocr.get(pn, {})
            exp_vu = int(info_o.get("PVF", 0) if tv == "SALIDA" else info_o.get("Costo c/IVA", 0))
            exp_tt = qv * exp_vu
            if int(round(to_float(rr.get("Valor Unit."), 0))) != exp_vu or int(round(to_float(rr.get("Total"), 0))) != exp_tt:
                recalc_ocr.at[ir, "Valor Unit."] = exp_vu
                recalc_ocr.at[ir, "Total"] = exp_tt
                need_ocr_refresh = True
        if need_ocr_refresh:
            st.session_state["ocr_rows"] = recalc_ocr
            st.session_state["ocr_ver"] = st.session_state.get("ocr_ver", 0) + 1
            st.rerun()

        valid_ocr = edited_ocr[edited_ocr["Producto"].notna() & (edited_ocr["Producto"].astype(str).str.strip() != "")]
        st.markdown(f"#### 💰 Total detectado: **${int(valid_ocr['Total'].sum()):,} COP** ({len(valid_ocr)} líneas válidas)")

        if st.button("💾 Guardar Todo el Día en Excel", type="primary", width="stretch"):
            if valid_ocr.empty:
                st.warning("⚠️ No hay filas con producto seleccionado para guardar.")
            else:
                try:
                    items_to_save = []
                    for _, rr in valid_ocr.iterrows():
                        items_to_save.append(
                            {
                                "fecha": ocr_fecha,
                                "tipo": str(rr["Tipo"] or ocr_tipo_def).strip().upper(),
                                "producto": str(rr["Producto"]).strip(),
                                "cantidad": int(round(to_float(rr["Cantidad"], 1))),
                                "cliente": str(rr.get("Nota / Cliente") or "").strip(),
                                "pago": str(rr.get("Medio de Pago") or ocr_pago_def).strip(),
                                "obs": f"Libreta: {rr.get('Texto Libreta', '')}".strip(),
                            }
                        )
                    filas = save_cart_movements(items_to_save)
                    st.session_state.pop("ocr_rows", None)
                    st.session_state.pop("hist_work", None)
                    st.session_state["hist_ver"] = st.session_state.get("hist_ver", 0) + 1
                    destino_txt = "Google Sheets" if is_cloud_mode() else "LISTADO_VENTA.xlsx"
                    st.session_state["popup_msg"] = f"¡Guardadas {len(filas)} líneas de la libreta en {destino_txt} (filas {filas[0]} a {filas[-1]})!"
                    st.rerun()
                except PermissionError:
                    st.error("❌ Cierra `LISTADO_VENTA.xlsx` en Excel para poder guardar.")
                except Exception as e:
                    st.error(f"❌ Error al guardar: {e}")
