using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LocalizacionColombia
{
    public class Factura_Traslado
    {
        public string SN { get; set; }
        public string fechaPed { get; set; }
        public string fechaCir { get; set; }
        public string nombreMed { get; set; }
        public string nombrePac { get; set; }
        public string clinica { get; set; }
        public string cedula { get; set; }
        public string aFactura { get; set; }
        public string ubicacion { get; set; }
        public string tipoDoc { get; set; }
        //Trasnferencia Devolucion
        public string transRef { get; set; }
        //Trasnferencia Original
        public string transOrg { get; set; }
        //Lineas
    }
    public class lineas_Factura
    {
        public string codigo { get; set; }
        public string descrip { get; set; }
        public Dictionary<string, string> serie { get; set; }
        public List<string> lote { get; set; }
        public List<string> fechaVcto { get; set; }
        public int numLinea { get; set; }
        public int cantidad { get; set; }
        public double precio { get; set; }
        public string impuesto { get; set; }
        public string almacen { get; set; }
        public string ciudad { get; set; }
        public string centroCostos { get; set; }
    }
}
