using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportExportFile.Models
{
    public class Quantity
    {
        public int id { get; set; }
        public int region_id { get; set; }
        public int company_id { get; set; }
        public int product_id { get; set; }
        public double quantity { get; set; }
        public DateTime create_time { get; set; }

    }
}