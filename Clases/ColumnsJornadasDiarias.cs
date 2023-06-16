using System;

namespace WestRockDataPonchesPRO.Clases
{
    public class ColumnsJornadasDiarias
    {
        public string UserId { get; set; }
        public string Empleado { get; set; }
        public string Horario { get; set; }
        public string Horario_Entrada { get; set; }
        public string Horario_Salida { get; set; }
        public DateTime Fecha { get; set; }
        public string Mark1 { get; set; }
        public string Mark2 { get; set; }
        public string Mark3 { get; set; }
        public string Mark4 { get; set; }
        public Int32 Ponches { get; set; }
        public Double Horas_Jornada { get; set; }
        public TimeSpan Horas_Extras { get; set; }
        public Double HorasExtras1 { get; set; }
        public Int32 Factor1 { get; set; }
        public Double SalarioHora { get; set; }
        public Double SalarioFraccion1 { get; set; }
        public Double MontoExtra1 { get; set; }
        public Double HorasExtras2 { get; set; }
        public Int32 Factor2 { get; set; }
        public Double SalarioFraccion2 { get; set; }
        public Double MontoExtra2 { get; set; }
        public Double TardanzaEntrada { get; set; }
    }
}
