using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Models
{
    public class ProjectMaterialTransformTransaction
    {
        
    }

    public class TransferLog
    {
        public string NewValueLocation { get; set; }
        public string ItemName { get; set; }
        public string Size { get; set; }
        public string NumberTransform { get; set; }
        public string OldValueLocation { get; set; }
    }
}
