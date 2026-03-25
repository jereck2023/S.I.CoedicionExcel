namespace CoedicionExcel.DTOs.Requests
{
    public class EliminarColumnaRequest
    {
        public int DocumentoId { get; set; }
        public string Columna { get; set; } = string.Empty;
    }
}