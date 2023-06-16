using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WestRockDataPonchesPRO.Clases
{
    public class ReporteHorasExtras
    {
        public int UserId { get; set; }
        public string UserName { get; set; }
        public string Jornada { get; set; }
        public string Departamento { get; set; }
        public DateTime FechaMarca { get; set; }
        public decimal HorasExtras { get; set; }
        public int Factor { get; set; }
        public double Salario { get; set; }
        public double SalarioFraccion { get; set; }
        public double Monto { get; set; }
    }
}
