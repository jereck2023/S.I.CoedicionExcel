namespace CoedicionExcel.Models
{
    public class FilaExcel
    {
        public int FilaId { get; set; }
        public int DocumentoId { get; set; }
        public string DatosJson { get; set; } = string.Empty;
        public int OrdenFila { get; set; }
        public bool Activa { get; set; } = true;
        public int VersionFila { get; set; } = 1;

        // Bloque 4 - estado de edición
        public bool EnEdicion { get; set; } = false;
        public string? EditadoPor { get; set; }
        public DateTime? FechaEdicion { get; set; }

        public DocumentoExcel? Documento { get; set; }
    }
}