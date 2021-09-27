using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LocalizacionColombia
{
    public static class DatosGlobalesFE
    {
        public static string URL { get; set; }
        public static string NIT { get; set; }
        public static string UsuarioWeb { get; set; }
        public static string PasswordWeb { get; set; }


        public static string IdentificadorFE { get; set; }
        public static string IdentificadorFEC { get; set; }
        public static string IdentificadorFEX { get; set; }


        public static int TimerReenvio { get; set; }
        public static int TimerEstado { get; set; }

        public static string localdecimal { get; set; }
        public static string localMillar { get; set; }
        public static string sapdecimal { get; set; }
        public static string sapMillar { get; set; }
    }
}
