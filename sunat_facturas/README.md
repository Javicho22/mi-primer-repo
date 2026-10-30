# Descarga de facturas desde SUNAT (API SIRE)

Programa en Python que descarga los comprobantes de **compras (RCE)** y **ventas (RVIE)**
registrados en SUNAT para uno o varios periodos. Con ellos arma un Excel listo
para la conciliación contable.

## 1. Requisitos previos en SUNAT

1. Ingresa a **SUNAT Operaciones en Línea** con el usuario principal.
2. Crea un **usuario secundario SOL** con perfil de acceso a **SIRE**.
3. Ve a *Empresas → Credenciales de API SUNAT* y registra una aplicación.
   Te dará un **client_id** y un **client_secret**.
4. Para la validación de comprobantes (opcional), la aplicación debe tener
   acceso a la API **"Consulta de validez de comprobantes"**.

## 2. Instalación

```bash
cd sunat_facturas
pip install -r requirements.txt
cp .env.example .env      # completa tus credenciales; .env NO se sube a GitHub
```

## 3. Uso

```bash
# Mes anterior, compras y ventas
python descargar_facturas.py

# Un periodo específico
python descargar_facturas.py --periodo 202609

# Rango de periodos, solo compras, validando cada CPE contra SUNAT
python descargar_facturas.py --desde 202601 --hasta 202609 --libro compras --validar
```

Resultado en `descargas/AAAAMM/`:

| Archivo | Contenido |
|---|---|
| `*.zip` | Archivo original de la propuesta SIRE (respaldo para fiscalización) |
| `comprobantes_<RUC>_<AAAAMM>.xlsx` | Hojas de detalle y resumen por tipo de comprobante (base imponible, IGV, no gravado, total) |

## 4. Uso contable

- **Crédito fiscal (arts. 18 y 19 de la LIGV):** concilia la hoja de compras con el
  Registro de Compras contable antes de aceptar la propuesta. La opción `--validar`
  marca los CPE **ANULADOS** o **NO EXISTE**. Esos comprobantes no sustentan crédito
  fiscal ni gasto deducible (art. 37 LIR, art. 44 inc. j).
- **Notas de crédito (07):** el resumen las muestra con signo negativo y las resta
  de la base imponible y del IGV, como ajuste del periodo.
- **Devengo (NIC 1 / Marco Conceptual):** el periodo SIRE corresponde a la fecha de
  emisión o anotación del CPE. Los comprobantes recibidos fuera de plazo deben
  analizarse por separado: su provisión contable puede ser de otro mes.
- **Ventas (NIIF 15):** compara el RVIE con los ingresos reconocidos. Las diferencias
  suelen venir de anticipos, ventas no facturadas o NC pendientes.
- **Conservación:** guarda los ZIP originales mientras la obligación tributaria
  no prescriba (art. 87 numeral 7 y art. 43 del Código Tributario).

## 5. Notas técnicas

- Las URLs de la API están al inicio de `descargar_facturas.py`. Si SUNAT publica
  una nueva versión del Manual SIRE, solo hay que actualizarlas ahí.
- SUNAT procesa las descargas por **ticket**. El programa consulta el estado cada
  15 segundos, hasta por 10 minutos.
