using PDF2XLS.Models;
using PDF2XLS.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
// ReSharper disable All

namespace PDF2XLS
{
    static class Program
    {
        static async Task Main(string[] args)
        {
            string pdfFile = Directory.GetCurrentDirectory() + @"\Files";
            DirectoryInfo folder = new DirectoryInfo(pdfFile);
            if (folder.GetFiles().Length > 0)
            {
                var citylist = ComHelper.GetCityTownList();
                List<Bsjh> bsjhs = new List<Bsjh>();
                #region PDF->XML
                Console.WriteLine("开始生成XML");
                foreach (FileInfo file in folder.GetFiles())
                {
                    if (file.Extension.ToUpper() == ".PDF")
                    {
                        var b = ComHelper.PdfToXMLAsFiles(file.FullName);
                        Console.WriteLine(b ? file.Name + "生成数据成功" : file.Name + "生成数据失败");
                    }
                }
                Console.WriteLine("生成XML结束");
                #endregion

                #region XML->LIST
                int index = 1;
                string yueri = DateTime.Now.AddDays(1).Date.ToString("MM-dd");
                string xq = ComHelper.GetWeek();
                string jihuafile = yueri + " 星期" + xq + ".xlsx";
                foreach (FileInfo file in folder.GetFiles())
                {
                    if (file.Extension.ToUpper() == ".XML")
                    {

                        Bsjh bsjh = new Bsjh
                        {
                            Id = index,
                            Riqi = yueri,
                            XingQi = xq,
                            IceCar = "Y"
                        };
                        index++;
                        List<string> ishave = new List<string>();
                        XmlDocument doc = new XmlDocument();
                        doc.Load(file.FullName);
                        XmlNode rootNode = doc.DocumentElement;
                        StringBuilder sb = new StringBuilder();
                        StringBuilder lxsb = new StringBuilder();
                        foreach (XmlNode node in rootNode.ChildNodes)
                        {
                            if (node.Name == "page")
                            {
                                foreach (XmlNode node2 in node.ChildNodes)
                                {
                                    if (node2.Name == "table")
                                    {
                                        foreach (XmlNode node3 in node2.ChildNodes)
                                        {
                                            if (node3.Name == "row")
                                            {
                                                foreach (XmlNode cellnode in node3.ChildNodes)
                                                {
                                                    if (cellnode.Name == "cell")
                                                    {
                                                        var va = cellnode.InnerText;
                                                        #region 线路号
                                                        if (!ishave.Contains("lxh") && va.Contains("百胜中国运输派车单"))
                                                        {
                                                            va = va.Replace("百胜中国运输派车单 派车单号: ", "");
                                                            var varr = va.Split('-');
                                                            if (varr.Length > 0)
                                                            {
                                                                ishave.Add("lxh");
                                                                int.TryParse(varr[0], out int xianluhao);
                                                                bsjh.LuXianHao = xianluhao;
                                                            }
                                                        }
                                                        #endregion

                                                        #region 箱数
                                                        if (!ishave.Contains("ss") && va.Contains("总件数"))
                                                        {
                                                            ishave.Add("ss");
                                                            double.TryParse(cellnode.NextSibling.InnerText, out double ss);
                                                            bsjh.XiangShu = ss;
                                                        }
                                                        #endregion

                                                        #region 立方数
                                                        if (!ishave.Contains("lfs") && va.Contains("总体积立方"))
                                                        {
                                                            ishave.Add("lfs");
                                                            double.TryParse(cellnode.NextSibling.InnerText, out double lfs);
                                                            bsjh.LiFangShu = lfs;
                                                        }
                                                        #endregion

                                                        #region 重量
                                                        if (!ishave.Contains("zl") && va.Contains("总重量KG"))
                                                        {
                                                            ishave.Add("zl");
                                                            double.TryParse(cellnode.NextSibling.InnerText, out double zl);
                                                            bsjh.ZhongLiang = zl;
                                                        }
                                                        #endregion

                                                        #region 公里数
                                                        if (!ishave.Contains("gls") && va.Contains("总里程"))
                                                        {
                                                            ishave.Add("gls");
                                                            double.TryParse(cellnode.NextSibling.InnerText, out double gls);
                                                            bsjh.GongLiShu = gls;
                                                        }
                                                        #endregion

                                                        #region 吨位
                                                        if (!ishave.Contains("dw") && va.Contains("车型:"))
                                                        {
                                                            var arry = cellnode.NextSibling.InnerText.Split("吨", StringSplitOptions.RemoveEmptyEntries);
                                                            if (arry.Length > 0)
                                                            {
                                                                ishave.Add("dw");
                                                                bsjh.DunWei = arry[0] + "T";
                                                            }
                                                        }
                                                        #endregion

                                                        #region 客户,路线,路线名称 [正则匹配] 
                                                        string rex = @"[\d]{2}\..*\.[\d]*"; // 匹配站点
                                                        var m = Regex.Match(va, rex);
                                                        // 正则匹配:03.鞍山云景.80131116 这种形式的  
                                                        if (m.Success)
                                                        {
                                                            var zdstr = m.Value;
                                                            #region 路线
                                                            if (zdstr.Contains("01.")) // 客户
                                                            {
                                                                ishave.Add("01");
                                                                string kehu = "肯德基";
                                                                if (zdstr.Contains("天津"))
                                                                {
                                                                    kehu = "百胜(天津)";
                                                                }
                                                                if (zdstr.Contains("北京"))
                                                                {
                                                                    kehu = "百胜(北京)";
                                                                }
                                                                bsjh.KeHu = kehu;
                                                            }
                                                            else
                                                            {
                                                                var arry = zdstr.Split(".", StringSplitOptions.RemoveEmptyEntries);
                                                                if (arry.Length >= 2 && !ishave.Contains(arry[0]))
                                                                {
                                                                    ishave.Add(arry[0]);
                                                                    string zm = arry[1].Trim();
                                                                    string fiststr = zm.Substring(0, 1);
                                                                    if (int.TryParse(fiststr, out int _))
                                                                    {
                                                                        zm = zm.Replace(fiststr, "");
                                                                    }
                                                                    string laststr = zm.Substring(zm.Length - 1, 1);
                                                                    if (int.TryParse(laststr, out int _))
                                                                    {
                                                                        zm = zm.Replace(laststr, "");
                                                                    }
                                                                    sb.Append("-" + zm);
                                                                }

                                                                #region 路线名称
                                                                foreach (var cityTown in citylist)
                                                                {
                                                                    var city = cityTown.CityName.Replace("市", "");
                                                                    var town = cityTown.TownName.Replace("市", "").Replace("区", "").Replace("自治县", "").Replace("县", "");
                                                                    if (city != "" && zdstr.Contains(city) && !ishave.Contains(city))
                                                                    {
                                                                        ishave.Add(city);
                                                                        lxsb.Append("/" + city);
                                                                    }
                                                                    if (town != "" && zdstr.Contains(town) && !ishave.Contains(town))
                                                                    {
                                                                        ishave.Add(town);
                                                                        lxsb.Append("/" + town);
                                                                    }
                                                                }
                                                                #endregion
                                                            }
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (sb.Length > 0)
                        {
                            sb.Remove(0, 1);
                            bsjh.LuXian = sb.ToString();
                        }
                        if (lxsb.Length > 0)
                        {
                            lxsb.Remove(0, 1);
                            bsjh.LuXianName = lxsb.ToString();
                        }
                        bsjhs.Add(bsjh);
                    }
                }
                #endregion

                XyunJh jh = new XyunJh { Bsjhs = bsjhs };
                await ComHelper.ExportByTemplate(Directory.GetCurrentDirectory() + @"\JiHua\" + jihuafile, jh);
                try
                {
                    ComHelper.DelectDir(pdfFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("删除文件异常:" + ex);
                }

                Console.WriteLine("ok");
            }
            else
            {
                Console.WriteLine("请在程序Files目录添加pdf文件");
            }
            Console.ReadKey();

            #region 测试数据
            //List<Bsjh> bsjhs = new List<Bsjh>
            //{
            //    new Bsjh{Id=1,KeHu="肯德基",Riqi="2-22",XingQi="六",LuXianHao=6001,LuXianName="长春",XiangShu=275.16
            //    ,LiFangShu=275.16,ZhongLiang=3123.73,IceCar="Y",DunWei="3T",LuXian="长春钜城-长春环球中心-PH长春环球中心-长春龙嘉-长春东岭街-长春大经路- 长春万科缤纷里-长春君子兰-长春新长春站-九台鹏宏"
            //    ,Driver="",CarNumber="",DaoDaTime="",GongLiShu=275},
            //    new Bsjh{Id=1,KeHu="肯德基",Riqi="2-22",XingQi="六",LuXianHao=6001,LuXianName="长春",XiangShu=275.16
            //    ,LiFangShu=275.16,ZhongLiang=3123.73,IceCar="Y",DunWei="3T",LuXian="长春钜城-长春环球中心-PH长春环球中心-长春龙嘉-长春东岭街-长春大经路- 长春万科缤纷里-长春君子兰-长春新长春站-九台鹏宏"
            //    ,Driver="",CarNumber="",DaoDaTime="",GongLiShu=275},
            //    new Bsjh{Id=1,KeHu="肯德基",Riqi="2-22",XingQi="六",LuXianHao=6001,LuXianName="长春",XiangShu=275.16
            //    ,LiFangShu=275.16,ZhongLiang=3123.73,IceCar="Y",DunWei="3T",LuXian="长春钜城-长春环球中心-PH长春环球中心-长春龙嘉-长春东岭街-长春大经路- 长春万科缤纷里-长春君子兰-长春新长春站-九台鹏宏"
            //    ,Driver="",CarNumber="",DaoDaTime="",GongLiShu=275},
            //    new Bsjh{Id=1,KeHu="肯德基",Riqi="2-22",XingQi="六",LuXianHao=6001,LuXianName="长春",XiangShu=275.16
            //    ,LiFangShu=275.16,ZhongLiang=3123.73,IceCar="Y",DunWei="3T",LuXian="长春钜城-长春环球中心-PH长春环球中心-长春龙嘉-长春东岭街-长春大经路- 长春万科缤纷里-长春君子兰-长春新长春站-九台鹏宏"
            //    ,Driver="",CarNumber="",DaoDaTime="",GongLiShu=275}
            //};
            #endregion

        }
    }
}
