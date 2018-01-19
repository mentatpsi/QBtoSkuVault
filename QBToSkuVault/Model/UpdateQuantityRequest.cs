using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBToSkuVault.Model
{
    public class UpdateQuantityRequest
    {
        public string Sku { get; set; }
        public string WarehouseId { get; set; }
        public string LocationCode { get; set; }
        public string Quantity { get; set; }
    }
}
