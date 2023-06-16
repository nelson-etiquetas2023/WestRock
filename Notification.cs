using System;

namespace WestRockDataPonchesPRO
{
    public class Notification
    {
        public string  Tipo { get; set; }
        public DateTime Fecha1 { get; set; }
        public DateTime Fecha2 { get; set; }
        public string Iduser { get; set; }
        public string Empleado { get; set; }
        public String Marcaje1 { get; set; }
        public String Marcaje2 { get; set; }
        public decimal Diffminutes { get; set; }
        public int Tardanza { get; set; }
        public string Device { get; set; }
        public string NameDevice { get; set; }
        public int Marcas { get; set; }

    }
}
