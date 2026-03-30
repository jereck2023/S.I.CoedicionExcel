using System;
using System.Collections.Generic;

namespace CoedicionExcel.Models
{
    public class DocumentoExcel
    {
        public int DocumentoId { get; set; }
        public string NombreArchivo { get; set; } = string.Empty;
        public DateTime FechaCarga { get; set; }
        public int Version { get; set; } = 1;

        // Bloque 4 - estado de edición de encabezados
        public bool EncabezadoEnEdicion { get; set; } = false;
        public string? EncabezadoEditadoPor { get; set; }
        public DateTime? FechaEdicionEncabezado { get; set; }

        public List<ColumnaExcel> Columnas { get; set; } = new();
        public List<FilaExcel> Filas { get; set; } = new();
    }
}