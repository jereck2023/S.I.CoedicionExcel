namespace CoedicionExcel.DTOs.Requests
{
    public class EstadoEdicionEncabezadoRequest
    {
        public int DocumentoId { get; set; }
        public string Usuario { get; set; } = string.Empty;
    }
}