using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBToSkuVault.Model
{
    public class CreateProductRequest
    {
        public string Brand { get; set; }
        public string Classification { get; set; }
        public string Sku { get; set; }
        public SupplierInfo SupplierInfo { get; set; }
    }
}
