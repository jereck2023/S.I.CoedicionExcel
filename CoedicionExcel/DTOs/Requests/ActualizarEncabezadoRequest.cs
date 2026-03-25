namespace CoedicionExcel.DTOs.Requests
{
    public class ActualizarEncabezadoRequest
    {
        public int DocumentoId { get; set; }
        public string Columna { get; set; } = string.Empty;
        public string NombreVisible { get; set; } = string.Empty;
    }
}