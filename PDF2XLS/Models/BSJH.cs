using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace PDF2XLS.Models
{
    /// <summary>
    /// 百盛计划model
    /// </summary>
    [ExcelExporter(Name = "运输部排车计划", TableStyle = "Light1", AutoFitAllColumn = false, MaxRowNumberOnASheet = 2)]
    public class Bsjh
    {
 
        [ExporterHeader(DisplayName = "序号")]
        public int Id { get; set; }
        
        [ExporterHeader(DisplayName = "客户")]
        public string KeHu { get; set; }
        
        [ExporterHeader(DisplayName = "日期")]
        public string Riqi { get; set; }
        
        [ExporterHeader(DisplayName = "星期")]
        public string XingQi { get; set; }
        
        [ExporterHeader(DisplayName = "线路号")]
        public string LuXianHao { get; set; }
        
        [ExporterHeader(DisplayName = "线路名称")]
        public string LuXianName { get; set; }
        
        [ExporterHeader(DisplayName = "箱数")]
        public string XiangShu { get; set; }
        
        [ExporterHeader(DisplayName = "立方数")]
        public string LiFangShu { get; set; }
        
        [ExporterHeader(DisplayName = "重量")]
        public string ZhongLiang { get; set; }
        
        [ExporterHeader(DisplayName = "冷藏车")]
        public string IceCar { get; set; }
        
        [ExporterHeader(DisplayName = "吨位")]
        public string DunWei { get; set; }
        
        [ExporterHeader(DisplayName = "线路")]
        public string LuXian { get; set; }
        
        [ExporterHeader(DisplayName = "司机")]
        public string Driver { get; set; }
       
        [ExporterHeader(DisplayName = "车牌号")]
        public string CarNumber { get; set; }
        
        [ExporterHeader(DisplayName = "到达仓库时间")]
        public string DaoDaTime { get; set; }

        [ExporterHeader(DisplayName = "公里数")]
        public string GongLiShu { get; set; }
    }
}
