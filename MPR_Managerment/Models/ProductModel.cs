using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Models
{
    public class ProductModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Des2 { get; set; }
        public string Code { get; set; }
        public string ProdMaterialCode { get; set; }
        public string PictureLink { get; set; }
        public byte[] Picture { get; set; } // Dạng mảng byte cho hình ảnh
        public string A_Thickness { get; set; }
        public string B_Depth { get; set; }
        public string C_Width { get; set; }
        public string D_Web { get; set; }
        public string E_Flag { get; set; }
        public string F_Length { get; set; }
        public string G_Weight { get; set; }
        public string UsedNote { get; set; }
        public int ProdOriginId { get; set; }
        public int ProdStandardId { get; set; }
        public int ProdMaterialCateId { get; set; }
        public int ProdMaterialId { get; set; }
        public int ProdMaterialDetailId { get; set; }
    }
}
