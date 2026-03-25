namespace CoedicionExcel.Models
{
    public class ColumnaExcel
    {
        public int ColumnaId { get; set; }
        public int DocumentoId { get; set; }
        public string ClaveColumna { get; set; } = string.Empty;
        public string NombreVisible { get; set; } = string.Empty;
        public int Orden { get; set; }
        public bool Activa { get; set; } = true;
        public DocumentoExcel? Documento { get; set; }
    }
}
