#!/usr/bin/env python3
"""
Descarga de comprobantes (facturas) desde SUNAT vía API SIRE.

Flujo:
  1. Obtiene token OAuth2 con las credenciales API generadas en Menú SOL.
  2. Solicita la exportación de la propuesta del Registro de Compras (RCE)
     y/o del Registro de Ventas (RVIE) para cada periodo.
  3. Consulta el ticket hasta que SUNAT termine de generar el archivo.
  4. Descarga el ZIP, lo descomprime y arma un Excel con el detalle de
     comprobantes y un resumen (base imponible, IGV, total) por tipo de CP.
  5. Opcionalmente valida cada comprobante de compras con la API de
     "Consulta de validez de CPE" (sustento del crédito fiscal, art. 18-19 LIGV).

Uso:
  python descargar_facturas.py --periodo 202609
  python descargar_facturas.py --desde 202601 --hasta 202609 --libro compras --validar
"""

import argparse
import csv
import io
import os
import sys
import time
import unicodedata
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# --- Endpoints SUNAT (verificar contra el Manual SIRE vigente) --------------
URL_TOKEN_SOL = "https://api-seguridad.sunat.gob.pe/v1/clientessol/{client_id}/oauth2/token/"
URL_TOKEN_EXTRANET = "https://api-seguridad.sunat.gob.pe/v1/clientesextranet/{client_id}/oauth2/token/"
SCOPE_SIRE = "https://api-sire.sunat.gob.pe"
SCOPE_VALIDEZ = "https://api.sunat.gob.pe/v1/contribuyente/contribuyentes"

BASE_SIRE = "https://api-sire.sunat.gob.pe/v1/contribuyente/migeigv/libros"
URL_VALIDEZ = "https://api.sunat.gob.pe/v1/contribuyente/contribuyentes/{ruc}/validarcomprobante"

LIBROS = {
    # Registro de Compras Electrónico (Formato 8.1)
    "compras": {
        "cod_libro": "080000",
        "exportar": BASE_SIRE + "/rce/propuesta/web/propuesta/{periodo}/exportacioncomprobantepropuesta",
        "params": {"codTipoArchivo": "0", "codOrigenEnvio": "2"},
    },
    # Registro de Ventas e Ingresos Electrónico (Formato 14.1)
    "ventas": {
        "cod_libro": "140000",
        "exportar": BASE_SIRE + "/rvie/propuesta/web/propuesta/{periodo}/exportapropuesta",
        "params": {"codTipoArchivo": "0"},
    },
}
URL_CONSULTA_TICKET = BASE_SIRE + "/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets"
URL_DESCARGA = BASE_SIRE + "/rvierce/gestionprocesosmasivos/web/masivo/archivoreporte"

ESTADOS_CPE = {
    "0": "NO EXISTE",
    "1": "ACEPTADO",
    "2": "ANULADO",
    "3": "AUTORIZADO",
    "4": "NO AUTORIZADO",
}
TIPOS_CP = {
    "01": "Factura",
    "03": "Boleta de venta",
    "07": "Nota de crédito",
    "08": "Nota de débito",
    "14": "Recibo servicios públicos",
}


# --- Configuración -----------------------------------------------------------
def cargar_env(ruta=".env"):
    """Carga variables KEY=VALUE de un archivo .env sin dependencias extra."""
    ruta = Path(ruta)
    if not ruta.exists():
        return
    for linea in ruta.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or "=" not in linea:
            continue
        clave, valor = linea.split("=", 1)
        os.environ.setdefault(clave.strip(), valor.strip().strip('"').strip("'"))


def requerir(nombre):
    valor = os.environ.get(nombre)
    if not valor:
        sys.exit(f"Falta la variable {nombre}. Configúrala en el archivo .env (ver .env.example).")
    return valor


# --- Autenticación -----------------------------------------------------------
def token_sire():
    client_id = requerir("SUNAT_CLIENT_ID")
    data = {
        "grant_type": "password",
        "scope": SCOPE_SIRE,
        "client_id": client_id,
        "client_secret": requerir("SUNAT_CLIENT_SECRET"),
        "username": requerir("SUNAT_RUC") + requerir("SUNAT_USUARIO_SOL"),
        "password": requerir("SUNAT_CLAVE_SOL"),
    }
    r = requests.post(URL_TOKEN_SOL.format(client_id=client_id), data=data, timeout=60)
    if r.status_code != 200:
        sys.exit(f"Error de autenticación SIRE ({r.status_code}): {r.text}")
    return r.json()["access_token"]


def token_validez():
    client_id = os.environ.get("SUNAT_VALIDEZ_CLIENT_ID") or requerir("SUNAT_CLIENT_ID")
    secret = os.environ.get("SUNAT_VALIDEZ_CLIENT_SECRET") or requerir("SUNAT_CLIENT_SECRET")
    data = {
        "grant_type": "client_credentials",
        "scope": SCOPE_VALIDEZ,
        "client_id": client_id,
        "client_secret": secret,
    }
    r = requests.post(URL_TOKEN_EXTRANET.format(client_id=client_id), data=data, timeout=60)
    if r.status_code != 200:
        sys.exit(f"Error de autenticación API de validez ({r.status_code}): {r.text}")
    return r.json()["access_token"]


# --- SIRE: exportar, consultar ticket, descargar ------------------------------
def solicitar_exportacion(sesion, libro, periodo):
    cfg = LIBROS[libro]
    r = sesion.get(cfg["exportar"].format(periodo=periodo), params=cfg["params"], timeout=120)
    r.raise_for_status()
    ticket = r.json().get("numTicket")
    if not ticket:
        raise RuntimeError(f"SUNAT no devolvió ticket: {r.text}")
    return ticket


def esperar_ticket(sesion, periodo, ticket, intentos=40, pausa=15):
    params = {"perIni": periodo, "perFin": periodo, "page": 1, "perPage": 20, "numTicket": ticket}
    for _ in range(intentos):
        r = sesion.get(URL_CONSULTA_TICKET, params=params, timeout=60)
        r.raise_for_status()
        registros = r.json().get("registros") or []
        if registros:
            reg = registros[0]
            estado = (reg.get("desEstadoProceso") or "").lower()
            if "terminado" in estado:
                return reg
            if "error" in estado or "rechaz" in estado:
                raise RuntimeError(f"Ticket {ticket} terminó con estado: {reg.get('desEstadoProceso')}")
        print(f"    ticket {ticket} en proceso, reintentando en {pausa}s...")
        time.sleep(pausa)
    raise TimeoutError(f"El ticket {ticket} no terminó a tiempo; vuelve a intentar más tarde.")


def descargar_archivos(sesion, libro, periodo, ticket, registro, carpeta):
    rutas = []
    archivos = registro.get("archivoReporte") or []
    for arch in archivos:
        nombre = arch.get("nomArchivoReporte")
        tipo = arch.get("codTipoAchivoReporte") or arch.get("codTipoArchivoReporte")
        params = {
            "nomArchivoReporte": nombre,
            "codTipoArchivoReporte": tipo,
            "codLibro": LIBROS[libro]["cod_libro"],
            "perTributario": periodo,
            "codProceso": registro.get("codProceso"),
            "numTicket": ticket,
        }
        r = sesion.get(URL_DESCARGA, params=params, timeout=300)
        r.raise_for_status()
        destino = carpeta / nombre
        destino.write_bytes(r.content)
        rutas.append(destino)
    return rutas


# --- Lectura del TXT de la propuesta ----------------------------------------
def normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    return "".join(c for c in texto.lower() if c.isalnum())


def leer_propuesta(ruta_zip):
    """Devuelve (encabezados, filas) del primer TXT/CSV dentro del ZIP."""
    with zipfile.ZipFile(ruta_zip) as zf:
        nombre = next(n for n in zf.namelist() if n.lower().endswith((".txt", ".csv")))
        contenido = zf.read(nombre)
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            texto = contenido.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    lector = csv.reader(io.StringIO(texto), delimiter="|")
    filas = [f for f in lector if any(c.strip() for c in f)]
    if not filas:
        return [], []
    encabezados = [c.strip() for c in filas[0]]
    return encabezados, [f[: len(encabezados)] for f in filas[1:]]


def buscar_columna(encabezados, *claves, excluir=()):
    """Primera columna cuyo nombre normalizado contiene todas las claves."""
    norm = [normalizar(e) for e in encabezados]
    for i, n in enumerate(norm):
        if all(k in n for k in claves) and not any(x in n for x in excluir):
            return i
    return None


def a_numero(valor):
    try:
        return float(str(valor).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


# --- Validación de comprobantes ------------------------------------------------
def validar_compras(encabezados, filas, ruc_consulta):
    col = {
        "ruc": buscar_columna(encabezados, "nrodoc", "identidad"),
        "tipo": buscar_columna(encabezados, "tipo", "cp", excluir=("ref", "mod", "identidad")),
        "serie": buscar_columna(encabezados, "serie", excluir=("ref", "mod", "dam")),
        "numero": buscar_columna(encabezados, "nrocp", excluir=("ref", "mod", "final"))
        or buscar_columna(encabezados, "numero", excluir=("ref", "mod", "final", "identidad")),
        "fecha": buscar_columna(encabezados, "fecha", "emision"),
        "total": buscar_columna(encabezados, "total", "cp"),
    }
    faltantes = [k for k, v in col.items() if v is None]
    if faltantes:
        print(f"    No se pudo validar: columnas no encontradas {faltantes}")
        return None

    sesion = requests.Session()
    sesion.headers["Authorization"] = f"Bearer {token_validez()}"
    url = URL_VALIDEZ.format(ruc=ruc_consulta)
    resultados = []
    for fila in filas:
        tipo = fila[col["tipo"]].strip().zfill(2)
        if tipo not in ("01", "03", "07", "08"):
            resultados.append("NO APLICA")
            continue
        cuerpo = {
            "numRuc": fila[col["ruc"]].strip(),
            "codComp": tipo,
            "numeroSerie": fila[col["serie"]].strip(),
            "numero": str(int(a_numero(fila[col["numero"]]))),
            "fechaEmision": fila[col["fecha"]].strip(),
            "monto": f"{a_numero(fila[col['total']]):.2f}",
        }
        try:
            r = sesion.post(url, json=cuerpo, timeout=30)
            datos = r.json().get("data") or {}
            resultados.append(ESTADOS_CPE.get(str(datos.get("estadoCp")), f"SIN RESPUESTA ({r.status_code})"))
        except (requests.RequestException, ValueError) as e:
            resultados.append(f"ERROR: {e}")
        time.sleep(0.2)
    return resultados


# --- Excel -------------------------------------------------------------------
ENCABEZADO = Font(bold=True, color="FFFFFF")
RELLENO = PatternFill("solid", fgColor="1F4E78")


def hoja_con_tabla(wb, titulo, encabezados, filas):
    ws = wb.create_sheet(titulo[:31])
    ws.append(encabezados)
    for celda in ws[1]:
        celda.font, celda.fill = ENCABEZADO, RELLENO
    for fila in filas:
        ws.append(fila)
    for i, enc in enumerate(encabezados, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(str(enc)) + 2))
    ws.freeze_panes = "A2"
    return ws


def resumen_por_tipo(encabezados, filas):
    """Resumen contable por tipo de comprobante: BI gravada, IGV, no gravado, total."""
    i_tipo = buscar_columna(encabezados, "tipo", "cp", excluir=("ref", "mod", "identidad"))
    i_bi = buscar_columna(encabezados, "bi", "grava")
    i_igv = buscar_columna(encabezados, "igv", excluir=("ng", "dng", "dgng"))
    i_nogr = buscar_columna(encabezados, "nogravad") or buscar_columna(encabezados, "valor", "adq", "ng")
    i_total = buscar_columna(encabezados, "total", "cp")
    if i_tipo is None or i_total is None:
        return None
    acum = defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0.0])
    for f in filas:
        tipo = f[i_tipo].strip().zfill(2)
        a = acum[tipo]
        a[0] += 1
        a[1] += a_numero(f[i_bi]) if i_bi is not None else 0
        a[2] += a_numero(f[i_igv]) if i_igv is not None else 0
        a[3] += a_numero(f[i_nogr]) if i_nogr is not None else 0
        a[4] += a_numero(f[i_total])
    filas_res = [
        [t, TIPOS_CP.get(t, "Otro"), *[round(x, 2) if isinstance(x, float) else x for x in v]]
        for t, v in sorted(acum.items())
    ]
    total = ["", "TOTAL"] + [round(sum(r[i] for r in filas_res), 2) for i in range(2, 7)]
    return filas_res + [total]


# --- Programa principal ------------------------------------------------------
def periodos_entre(desde, hasta):
    a, m = int(desde[:4]), int(desde[4:])
    while f"{a}{m:02d}" <= hasta:
        yield f"{a}{m:02d}"
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)


def main():
    cargar_env(Path(__file__).with_name(".env"))
    hoy = date.today()
    mes_anterior = f"{hoy.year - (hoy.month == 1)}{(hoy.month - 2) % 12 + 1:02d}"

    p = argparse.ArgumentParser(description="Descarga tus comprobantes desde SUNAT (API SIRE).")
    p.add_argument("--periodo", help="Periodo AAAAMM (por defecto, el mes anterior)")
    p.add_argument("--desde", help="Periodo inicial AAAAMM")
    p.add_argument("--hasta", help="Periodo final AAAAMM")
    p.add_argument("--libro", choices=["compras", "ventas", "ambos"], default="ambos")
    p.add_argument("--validar", action="store_true",
                   help="Validar cada comprobante de compras con la API de validez de CPE")
    p.add_argument("--salida", default="descargas", help="Carpeta de salida")
    args = p.parse_args()

    if args.desde:
        periodos = list(periodos_entre(args.desde, args.hasta or args.desde))
    else:
        periodos = [args.periodo or mes_anterior]
    libros = ["compras", "ventas"] if args.libro == "ambos" else [args.libro]
    ruc = requerir("SUNAT_RUC")

    sesion = requests.Session()
    sesion.headers["Authorization"] = f"Bearer {token_sire()}"
    print(f"Autenticado en SIRE para RUC {ruc}.")

    for periodo in periodos:
        carpeta = Path(args.salida) / periodo
        carpeta.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        wb.remove(wb.active)
        for libro in libros:
            print(f"[{periodo}] {libro}: solicitando propuesta...")
            try:
                ticket = solicitar_exportacion(sesion, libro, periodo)
                registro = esperar_ticket(sesion, periodo, ticket)
                zips = descargar_archivos(sesion, libro, periodo, ticket, registro, carpeta)
            except (requests.HTTPError, RuntimeError, TimeoutError) as e:
                print(f"    No se pudo descargar {libro} {periodo}: {e}")
                continue
            for z in zips:
                if not zipfile.is_zipfile(z):
                    continue
                encabezados, filas = leer_propuesta(z)
                print(f"    {len(filas)} comprobantes en {z.name}")
                if args.validar and libro == "compras" and filas:
                    estados = validar_compras(encabezados, filas, ruc)
                    if estados:
                        encabezados = encabezados + ["Estado CPE (SUNAT)"]
                        filas = [f + [e] for f, e in zip(filas, estados)]
                hoja_con_tabla(wb, f"{libro.capitalize()} detalle", encabezados, filas)
                resumen = resumen_por_tipo(encabezados, filas)
                if resumen:
                    hoja_con_tabla(
                        wb, f"{libro.capitalize()} resumen",
                        ["Tipo CP", "Descripción", "Cantidad", "Base imponible gravada",
                         "IGV", "No gravado", "Importe total"],
                        resumen,
                    )
        if wb.sheetnames:
            excel = carpeta / f"comprobantes_{ruc}_{periodo}.xlsx"
            wb.save(excel)
            print(f"[{periodo}] Excel generado: {excel}")


if __name__ == "__main__":
    main()
