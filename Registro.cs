using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace WestRockDataPonchesPRO
{
    [Serializable]
    public class Registro
    {
        public int NumberRecord { get; set; }
        public string UserID { get; set; }
        public string UserName { get; set; }
        public DateTime FechaHora_Marca { get; set; }
        public string Device { get; set; }
        public string NameDevice { get; set; }
        public string DateRegistro { get; set; }
        public string HourRegistro { get; set; }
        public string Reference { get; set; }
        public string HoraMark { get; set; }
        public string Jornada { get; set; }
        public string Departamento { get; set; }
        public string ShiftId { get; set; }
        public string shiftname { get; set; }
        public string type_Shift { get; set; }

    }
}
