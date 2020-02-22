using PDF2XLS.Models;
using PDF2XLS.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace PDF2XLS
{
    static class Program
    {
        static async Task Main(string[] args)
        {
            string pdfFile = Directory.GetCurrentDirectory() + @"\Files";
            //DirectoryInfo folder = new DirectoryInfo(pdfFile);
            //foreach (FileInfo file in folder.GetFiles())
            //{
            //    if (file.Extension.ToUpper() == ".PDF")
            //    {
            //        var b = ComHelper.PdfToXMLAsFiles(file.FullName);
            //        Console.WriteLine(b ? file.Name + "生成数据成功" : file.Name + "生成数据失败");
            //    }
            //}
            List<Bsjh> bsjhs = new List<Bsjh>
            {
                new Bsjh{Id=1,KeHu="肯德基",Riqi="2-22",XingQi="六",LuXianHao="6001",LuXianName="长春",XiangShu="275.16"
                ,LiFangShu="275.16",ZhongLiang="3123.73",IceCar="Y",DunWei="3T",LuXian="长春钜城-长春环球中心-PH长春环球中心-长春龙嘉-长春东岭街-长春大经路- 长春万科缤纷里-长春君子兰-长春新长春站-九台鹏宏"
                ,Driver="",CarNumber="",DaoDaTime="",GongLiShu=""
                }
            };
            await ComHelper.Export(Directory.GetCurrentDirectory()+@"\jihua.xlsx", bsjhs);
            Console.WriteLine("ok");
            Console.ReadKey();
        }
    }
}
