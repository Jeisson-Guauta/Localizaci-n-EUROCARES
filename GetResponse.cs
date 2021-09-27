using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LocalizacionColombia
{
    public class GetFieldID
    {

        public string Name { set; get; }
        public string Type { get; set; }
        public string Size { get; set; }
        public int FieldID { get; set; }
    }
    public class GetQueryCategoryID
    {
        public int Code { get; set; }
    }
    public class GetQueryID
    {
        public int InternalKey { get; set; }
    }
    public class ODataResponse<T>
    {
        public List<T> Value { get; set; }
    }
}
