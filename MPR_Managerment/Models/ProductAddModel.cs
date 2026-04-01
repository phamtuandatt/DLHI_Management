using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Models
{
    public class ProductAddModel
    {
        public int Id { get; set; }
        public string Code { get; set; }
    }

    public class MaterialCategoris
    {
        public int Cate_Id { get; set; }
        public string Cate_Name { get; set; }
    }

    public class Materials
    {
        public int Material_Id { get; set; }
        public string Materia_Code { get; set; }
        public string Material_Name { get; set; }
        public string Specifications { get; set; }
        public DateTime CreatedDate { get; set; }

        public int Cate_Id { get; set; }

    }

    public class Material_Detail
    {
        public int Material_Detail_Id { get; set; }
        public string Material_Detail_Number { get; set; }
        public string Material_Detail_Name{ get; set; }
        public string Material_Detail_Code { get; set; }
        public string Item_Code_Existed { get; set; }
    }
}
