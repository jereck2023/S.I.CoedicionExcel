namespace CoedicionExcel.DTOs.Requests
{
    public class ActualizarCeldaRequest
    {
        public int FilaId { get; set; }
        public string Columna { get; set; } = string.Empty;
        public string Valor { get; set; } = string.Empty;
    }
}