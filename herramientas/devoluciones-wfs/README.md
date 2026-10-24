# Cuadra — control para vendedores WFS

Herramienta para cruzar los reportes de Walmart WFS y ver cuánto dinero hay en devoluciones, qué nunca llegó y qué reclamar.

## Cómo usarla

Abre `index.html` en el navegador y carga los reportes desde la pestaña **Cargar reportes**:

| Reporte | Dónde está en Seller Center | Qué aporta |
|---|---|---|
| Customer Returns (.xlsx) | WFS → Reports → Customer Returns | Qué se devolvió, quién tiene la culpa, adónde fue cada producto |
| Return to Seller (.csv) | WFS → Reports → Return to Seller | Lo que Walmart dice que te mandó de vuelta, con tracking |
| Settlement (.csv) | WFS → Reports → Settlement | Cargos y reembolsos de WFS |

Si abres el archivo directamente, los datos se guardan solo en ese navegador. La versión publicada en claude.ai los guarda en la nube.

## Qué hace

- Clasifica cada devolución: en camino, nunca llegó a Walmart, enviada a ti (por confirmar), te llegó / no te llegó, la cubre Walmart, Walmart te la pagó, desechada.
- Une cada devolución con la caja de FedEx en la que Walmart dice que te la mandó (por tracking + GTIN y, si no hay tracking, por fecha + GTIN).
- Detecta en el settlement cobros repetidos, cargos de devolución en devoluciones que son culpa de WFS, inventario "perdido" que luego te cobraron como "encontrado" y reembolsos de devoluciones perdidas.
- Arma el texto en inglés para abrir el caso en Seller Center.

Los reportes no se guardan en este repositorio porque tienen datos de clientes.
