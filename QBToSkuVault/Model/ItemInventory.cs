using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBToSkuVault.Model
{
    public class ItemInventory
    {
        public string ListID { get; set; }
        public string FullName { get; set; }
        public string QuantityOnHand { get; set; }
        public string ManufacturerPartNumber { get; set; }
    }
}
