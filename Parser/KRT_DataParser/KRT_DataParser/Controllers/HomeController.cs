using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Newtonsoft;
using Newtonsoft.Json;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using KRT_DataParser.Models;

namespace KRT_DataParser.Controllers
{
    public class HomeController : Controller
    {
        private List<Data> datas = new List<Data>();

        public ActionResult Index()
        {
            ParseOld(Server.MapPath("/Excels/Old/15-97.xls"), 97, 1);
            ParseOld(Server.MapPath("/Excels/Old/15-98.xls"), 98, 1);
            ParseOld(Server.MapPath("/Excels/Old/15-99.xls"), 99, 0);

            Parse(Server.MapPath("/Excels/New/2529-10-15-1000101.xls"), 100, 1);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000201.xls"), 100, 2);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000301.xls"), 100, 3);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000401.xls"), 100, 4);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000501.xls"), 100, 5);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000601.xls"), 100, 6);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000701.xls"), 100, 7);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1000801.xls"), 100, 8);
            //Parse(Server.MapPath("/Excels/New/2529-10-15-1000901.xls"), 100, 9); 暫時找不到
            Parse(Server.MapPath("/Excels/New/2529-10-15-1001001.xls"), 100, 10);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1001101.xls"), 100, 11);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1001201.xls"), 100, 12);

            Parse(Server.MapPath("/Excels/New/2529-10-15-1010101.xls"), 101, 1);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010201.xls"), 101, 2);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010301.xls"), 101, 3);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010401.xls"), 101, 4);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010501.xls"), 101, 5);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010601.xls"), 101, 6);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010701.xls"), 101, 7);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010801.xls"), 101, 8);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1010901.xls"), 101, 9);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1011001.xls"), 101, 10);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1011101.xls"), 101, 11);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1011201.xls"), 101, 12);

            Parse(Server.MapPath("/Excels/New/2529-10-15-1020101.xls"), 102, 1);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020201.xls"), 102, 2);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020301.xls"), 102, 3);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020401.xls"), 102, 4);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020501.xls"), 102, 5);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020601.xls"), 102, 6);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020701.xls"), 102, 7);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020801.xls"), 102, 8);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1020901.xls"), 102, 9);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1021001.xls"), 102, 10);
            Parse(Server.MapPath("/Excels/New/2529-10-15-1021101.xls"), 102, 11);

            string json = JsonConvert.SerializeObject(datas);

            return Content(json);
        }

        private void ParseOld(string url, int year, int vShift)
        {
            FileStream fs = new FileStream(url, FileMode.Open, FileAccess.Read);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);

            for (int s = 0; s < workbook.NumberOfSheets; s++)
            {
                HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(s);
                List<Station> stations = new List<Station>();

                for (int i = 6 + vShift; i <= sheet.LastRowNum; i++)
                {
                    HSSFRow row = (HSSFRow)sheet.GetRow(i);

                    if (row.GetCell(0) != null
                        && row.GetCell(0).ToString().Trim() != "")
                    {
                        //紅線
                        stations.Add(new Station()
                        {
                            ID = parseName(row.GetCell(0).StringCellValue, "紅線"),
                            In = Convert.ToInt32(row.GetCell(1).NumericCellValue),
                            Out = Convert.ToInt32(row.GetCell(1).NumericCellValue)
                        });
                        //橘線
                        if (row.GetCell(4) != null
                            && row.GetCell(4).ToString().Trim() != "")
                        {
                            stations.Add(new Station()
                            {
                                ID = parseName(row.GetCell(4).ToString(), "橘線"),
                                In = Convert.ToInt32(row.GetCell(5).NumericCellValue),
                                Out = Convert.ToInt32(row.GetCell(6).NumericCellValue)
                            });
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                //parse to json
                if (sheet.SheetName.Length >= 4)
                {
                    datas.Add(new Data()
                    {
                        Year = year,
                        Month = Convert.ToInt32(sheet.SheetName.ToString().Substring(2, 2)),
                        Stations = stations
                    });
                }
            }
            fs.Close();
        }

        private void Parse(string url, int year, int month)
        {
            FileStream fs = new FileStream(url, FileMode.Open, FileAccess.Read);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);
            HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(0);
            List<Station> stations = new List<Station>();

            for (int i = 6; i <= sheet.LastRowNum; i++)
            {
                HSSFRow row = (HSSFRow)sheet.GetRow(i);
                if (row.GetCell(0) != null
                    && row.GetCell(0).ToString().Trim() != ""
                    && row.GetCell(0).ToString().Trim() != "填表")
                {
                    //紅線
                    stations.Add(new Station()
                    {
                        ID = parseName(row.GetCell(0).StringCellValue, "紅線"),
                        In = Convert.ToInt32(row.GetCell(1).NumericCellValue),
                        Out = Convert.ToInt32(row.GetCell(2).NumericCellValue)
                    });
                    //橘線
                    if (row.GetCell(4) != null
                        && row.GetCell(4).ToString().Trim() != "")
                    {
                        stations.Add(new Station()
                        {
                            ID = parseName(row.GetCell(4).ToString(), "橘線"),
                            In = Convert.ToInt32(row.GetCell(5).NumericCellValue),
                            Out = Convert.ToInt32(row.GetCell(6).NumericCellValue)
                        });
                    }
                }
                else
                {
                    break;
                }
            }

            //parse to json
            datas.Add(new Data()
            {
                Year = year,
                Month = month,
                Stations = stations
            });

            fs.Close();
        }

        private string parseName(string name, string prefix)
        {
            string newName = name.Split(' ')[0];

            if (name == "團體票"
                || name == "其他")
            {
                newName = prefix + name;
            }

            return newName;
        }
    }
}
