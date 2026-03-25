namespace CoedicionExcel.DTOs.Requests
{
    public class AgregarColumnaRequest
    {
        public int DocumentoId { get; set; }
        public string Nombre { get; set; } = string.Empty;
    }
}