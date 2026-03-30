namespace CoedicionExcel.DTOs.Requests
{
    public class EstadoEdicionFilaRequest
    {
        public int FilaId { get; set; }
        public string Usuario { get; set; } = string.Empty;
    }
}