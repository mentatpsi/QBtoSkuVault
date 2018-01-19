using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBToSkuVault.Model
{
    public class CreateProductsRequest
    {
        public List<CreateProductRequest> Items { get; set; }
        public string TenantToken { get; set; }
        public string UserToken { get; set; }
    }
}
