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
        public DocumentoExcel? Documento { get; set; }
    }
}