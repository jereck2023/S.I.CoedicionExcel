using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using CoedicionExcel.Data;
using CoedicionExcel.Models;
using CoedicionExcel.DTOs.Requests;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;

namespace CoedicionExcel.Controllers
{
    public class ExcelController : Controller
    {
        private readonly AppDbContext _context;

        public ExcelController(AppDbContext context)
        {
            _context = context;
        }

        public IActionResult Index()
        {
            return View();
        }

        // GET
        public IActionResult Subir()
        {
            return View();
        }

        // POST
        [HttpPost]
        public async Task<IActionResult> Subir(IFormFile archivoExcel)
        {
            if (archivoExcel == null || archivoExcel.Length == 0)
                return BadRequest("Archivo inválido");

            using var stream = new MemoryStream();
            await archivoExcel.CopyToAsync(stream);

            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.First();

            var filas = worksheet.RangeUsed().RowsUsed().ToList();

            if (filas.Count < 2)
                return BadRequest("El Excel no tiene datos suficientes");

            //1. Crear documento
            var documento = new DocumentoExcel
            {
                NombreArchivo = archivoExcel.FileName,
                FechaCarga = DateTime.Now
            };

            _context.DocumentosExcel.Add(documento);
            await _context.SaveChangesAsync();

            //2. Leer encabezados
            var encabezados = filas[0].Cells().Select((c, i) => new ColumnaExcel
            {
                DocumentoId = documento.DocumentoId,
                ClaveColumna = $"col_{i + 1}",
                NombreVisible = c.GetValue<string>(),
                Orden = i + 1
            }).ToList();

            _context.ColumnasExcel.AddRange(encabezados);
            await _context.SaveChangesAsync();

            //3. Leer filas
            var filasData = new List<FilaExcel>();

            for (int i = 1; i < filas.Count; i++)
            {
                var fila = filas[i];
                var dict = new Dictionary<string, object>();

                for (int j = 0; j < encabezados.Count; j++)
                {
                    var valor = fila.Cell(j + 1).GetValue<string>();
                    dict[$"col_{j + 1}"] = valor;
                }

                filasData.Add(new FilaExcel
                {
                    DocumentoId = documento.DocumentoId,
                    DatosJson = JsonSerializer.Serialize(dict),
                    OrdenFila = i
                });
            }

            _context.FilasExcel.AddRange(filasData);
            await _context.SaveChangesAsync();

            return RedirectToAction("Tabla", new { id = documento.DocumentoId });
        }

        public async Task<IActionResult> Tabla(int id)
        {
            var documento = await _context.DocumentosExcel
                .FirstOrDefaultAsync(d => d.DocumentoId == id);

            if (documento == null)
                return NotFound();

            var columnas = await _context.ColumnasExcel
                .Where(c => c.DocumentoId == id && c.Activa)
                .OrderBy(c => c.Orden)
                .ToListAsync();

            var filasBd = await _context.FilasExcel
                .Where(f => f.DocumentoId == id && f.Activa)
                .OrderBy(f => f.OrdenFila)
                .ToListAsync();

            var filas = new List<Dictionary<string, object>>();

            foreach (var filaBd in filasBd)
            {
                var datos = JsonSerializer.Deserialize<Dictionary<string, object>>(filaBd.DatosJson)
                            ?? new Dictionary<string, object>();

                datos["filaId"] = filaBd.FilaId;
                filas.Add(datos);
            }

            ViewBag.DocumentoId = id;
            ViewBag.NombreArchivo = documento.NombreArchivo;
            ViewBag.Columnas = columnas;
            ViewBag.Filas = filas;

            return View();
        }

        //INICIO DE METODOS

        //Metodo para crear Endpoint
        [HttpPost]
        public async Task<IActionResult> ActualizarCelda([FromBody] ActualizarCeldaRequest request)
        {
            Console.WriteLine("=== ACTUALIZAR CELDA ===");

            if (request == null)
                return BadRequest("Request nulo");

            Console.WriteLine($"FilaId: {request.FilaId}");
            Console.WriteLine($"Columna: {request.Columna}");
            Console.WriteLine($"Valor: {request.Valor}");

            var fila = await _context.FilasExcel.FindAsync(request.FilaId);

            if (fila == null)
                return NotFound("Fila no encontrada");

            var datos = JsonSerializer.Deserialize<Dictionary<string, string>>(fila.DatosJson)
                        ?? new Dictionary<string, string>();

            datos[request.Columna] = request.Valor ?? "";

            fila.DatosJson = JsonSerializer.Serialize(datos);

            await _context.SaveChangesAsync();

            Console.WriteLine("Guardado correctamente en BD");

            return Ok("Guardado correctamente");
        }

        [HttpPost]
        public async Task<IActionResult> ActualizarEncabezado([FromBody] ActualizarEncabezadoRequest request)
        {
            if (request == null)
                return BadRequest("Request nulo");

            var columna = await _context.ColumnasExcel
                .FirstOrDefaultAsync(c => c.DocumentoId == request.DocumentoId
                                       && c.ClaveColumna == request.Columna);

            if (columna == null)
                return NotFound("Columna no encontrada");

            columna.NombreVisible = request.NombreVisible ?? "";

            await _context.SaveChangesAsync();

            return Ok("Encabezado actualizado");
        }

        [HttpPost]
        public async Task<IActionResult> AgregarFila([FromBody] AgregarFilaRequest request)
        {
            if (request == null || request.DocumentoId <= 0)
                return BadRequest("Documento inválido");

            var columnas = await _context.ColumnasExcel
                .Where(c => c.DocumentoId == request.DocumentoId && c.Activa)
                .OrderBy(c => c.Orden)
                .ToListAsync();

            if (!columnas.Any())
                return BadRequest("El documento no tiene columnas");

            var ultimaFila = await _context.FilasExcel
                .Where(f => f.DocumentoId == request.DocumentoId)
                .OrderByDescending(f => f.OrdenFila)
                .FirstOrDefaultAsync();

            int nuevoOrden = (ultimaFila?.OrdenFila ?? 0) + 1;

            var datos = new Dictionary<string, string>();

            foreach (var col in columnas)
            {
                datos[col.ClaveColumna] = "";
            }

            var nuevaFila = new FilaExcel
            {
                DocumentoId = request.DocumentoId,
                DatosJson = JsonSerializer.Serialize(datos),
                OrdenFila = nuevoOrden,
                Activa = true
            };

            _context.FilasExcel.Add(nuevaFila);
            await _context.SaveChangesAsync();

            var respuesta = new Dictionary<string, object>
            {
                ["filaId"] = nuevaFila.FilaId
            };

            foreach (var item in datos)
            {
                respuesta[item.Key] = item.Value;
            }

            return Json(respuesta);
        }

        [HttpPost]
        public async Task<IActionResult> EliminarFila([FromBody] EliminarFilaRequest request)
        {
            if (request == null || request.FilaId <= 0)
                return BadRequest("Fila inválida");

            var fila = await _context.FilasExcel.FindAsync(request.FilaId);

            if (fila == null)
                return NotFound("Fila no encontrada");

            fila.Activa = false;

            await _context.SaveChangesAsync();

            return Ok("Fila eliminada");
        }

        //Metodo para agregar columna
        [HttpPost]
        public async Task<IActionResult> AgregarColumna([FromBody] AgregarColumnaRequest request)
        {
            if (request == null || request.DocumentoId <= 0)
                return BadRequest("Documento inválido");

            var columnas = await _context.ColumnasExcel
                .Where(c => c.DocumentoId == request.DocumentoId)
                .ToListAsync();

            int siguienteNumero = columnas.Count + 1;
            string nuevaClave = $"col_{siguienteNumero}";

            int nuevoOrden = columnas.Any() ? columnas.Max(c => c.Orden) + 1 : 1;

            var nuevaColumna = new ColumnaExcel
            {
                DocumentoId = request.DocumentoId,
                ClaveColumna = nuevaClave,
                NombreVisible = request.Nombre,
                Orden = nuevoOrden,
                Activa = true
            };

            _context.ColumnasExcel.Add(nuevaColumna);

            //Actualizar todas las filas
            var filas = await _context.FilasExcel
                .Where(f => f.DocumentoId == request.DocumentoId && f.Activa)
                .ToListAsync();

            foreach (var fila in filas)
            {
                var datos = JsonSerializer.Deserialize<Dictionary<string, string>>(fila.DatosJson)
                            ?? new Dictionary<string, string>();

                datos[nuevaClave] = "";

                fila.DatosJson = JsonSerializer.Serialize(datos);
            }

            await _context.SaveChangesAsync();

            return Json(new
            {
                clave = nuevaClave,
                nombre = request.Nombre
            });
        }

        //Metodo para eliminar columnas
        [HttpPost]
        public async Task<IActionResult> EliminarColumna([FromBody] EliminarColumnaRequest request)
        {
            if (request == null || request.DocumentoId <= 0 || string.IsNullOrWhiteSpace(request.Columna))
                return BadRequest("Datos inválidos");

            var columna = await _context.ColumnasExcel
                .FirstOrDefaultAsync(c => c.DocumentoId == request.DocumentoId
                                       && c.ClaveColumna == request.Columna
                                       && c.Activa);

            if (columna == null)
                return NotFound("Columna no encontrada");

            columna.Activa = false;

            await _context.SaveChangesAsync();

            return Ok("Columna eliminada");
        }

        //Metodo para descargar Excel
        [HttpGet]
        public async Task<IActionResult> DescargarExcel(int documentoId)
        {
            var documento = await _context.DocumentosExcel
                .FirstOrDefaultAsync(d => d.DocumentoId == documentoId);

            if (documento == null)
                return NotFound();

            var columnas = await _context.ColumnasExcel
                .Where(c => c.DocumentoId == documentoId && c.Activa)
                .OrderBy(c => c.Orden)
                .ToListAsync();

            var filas = await _context.FilasExcel
                .Where(f => f.DocumentoId == documentoId && f.Activa)
                .OrderBy(f => f.OrdenFila)
                .ToListAsync();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Datos");

            //ENCABEZADOS
            for (int i = 0; i < columnas.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = columnas[i].NombreVisible;
            }

            //FILAS
            for (int i = 0; i < filas.Count; i++)
            {
                var datos = JsonSerializer.Deserialize<Dictionary<string, string>>(filas[i].DatosJson)
                            ?? new Dictionary<string, string>();

                for (int j = 0; j < columnas.Count; j++)
                {
                    var clave = columnas[j].ClaveColumna;

                    datos.TryGetValue(clave, out var valor);

                    worksheet.Cell(i + 2, j + 1).Value = valor ?? "";
                }
            }

            //TABLA EN EXCEL DE ESCRITORIO
            if (columnas.Count > 0)
            {
                int totalFilas = filas.Count + 1; //+1 por encabezado
                int totalColumnas = columnas.Count;

                var rango = worksheet.Range(1, 1, totalFilas, totalColumnas);
                var tabla = rango.CreateTable();

                tabla.Theme = XLTableTheme.TableStyleMedium2;
            } worksheet.Columns().AdjustToContents(); //Ajustar ancho

            //FIN DE LOS METODOS

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            string nombreArchivo = $"Exportado_{documentoId}.xlsx";

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo
            );
        }
    }
}