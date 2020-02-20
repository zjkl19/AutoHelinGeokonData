using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace AutoHelinGeokonData
{
    public class Program
    {
        public class DataIndex
        {
            /// <summary>
            /// 名称
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 列索引
            /// </summary>
            public int ColIndex { get; set; }

            public decimal MaxValue { get; set; }
            public decimal MinValue { get; set; }
            public decimal AverageValue { get; set; }

            public decimal BaseValue { get; set; } = 0.0m;

        }

        public class MonitorData
        {
            public string Name { get; set; }
            public decimal MaxValue { get; set; }
            public decimal MinValue { get; set; }
            public decimal AverageValue { get; set; }
        }
        static void Main(string[] args)
        {
            decimal errorValue = 0.00m;    //excel数据读取出错时返回的数值
            var dataFile = new FileInfo("设备监测数据.xlsx");
            var sheetName = "传感器监测数据报表";

            //静力水准仪基准数值（J8-1~J8-4）

            //定义
            var baseValue = new Dictionary<string, decimal>
            {
                //添加元素
                { "J8-1-位移(mm)", 49.11m },
                { "J8-2-位移(mm)", -44.52m },
                { "J8-3-位移(mm)", 56.88m },
                { "J8-4-位移(mm)", 57.16m }
            };
            //decimal[] baseValue = { 49.11m, -44.52m, 56.88m, 57.16m };    


            var inputOutputFile = new FileInfo($"输入输出.xlsx");
            try
            {

                using ExcelPackage excelPackage = new ExcelPackage(inputOutputFile)
                    , dataPackage = new ExcelPackage(dataFile);
                // 从第1列扫描到最后一列
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];
                int currCol = 1;
                string currChannel = string.Empty; string currValue = string.Empty;

                var inputList = new List<DataIndex>();    //原始输入
                var dataList = new List<MonitorData>();
                //用yield return 重构
                try
                {
                    while (!string.IsNullOrWhiteSpace(worksheet.Cells[1, currCol]?.Value?.ToString() ?? string.Empty))
                    {
                        currChannel = worksheet.Cells[1, currCol].Value.ToString();
                        if (currChannel.Substring(0, 1) == "S")    //裂缝
                        {
                            inputList.Add(new DataIndex { Name = $"{currChannel}-微应变(£)", ColIndex = currCol });
                        }
                        else if (currChannel.Substring(0, 1) == "J")
                        {
                            inputList.Add(new DataIndex { Name = $"{currChannel}-位移(mm)", ColIndex = currCol });
                        }
                        else
                        {
                            inputList.Add(new DataIndex { Name = currChannel, ColIndex = currCol });
                        }

                        //if (currChannel.Substring(0) == "S")    //裂缝
                        //{

                        //}

                        //ExcelWorksheet dataWorksheet = package.Workbook.Worksheets[sheetName];
                        ////第1行第2列
                        //var test = worksheet.Cells[1, 2]?.Value.ToString() ?? string.Empty;
                        //Console.WriteLine(test);
                        ////Todo:试试 Decimal.TryParse
                        //Console.WriteLine(Decimal.Parse(worksheet.Cells[688, 2]?.Value.ToString() ?? string.Empty));
                        currCol++;
                    }

                    decimal temp;
                    for (int i = 0; i < inputList.Count; i++)
                    {
                        baseValue.TryGetValue(inputList[i].Name, out temp);
                        inputList[i].BaseValue = temp;
                    }

                    var dataWorksheet = dataPackage.Workbook.Worksheets[sheetName];

                    //获取最大值、最小值、平均值所在行
                    int maxValueRow = 0;
                    int minValueRow = 0;
                    int averageValueRow = 0;

                    int currRow = 2;
                    while (!string.IsNullOrWhiteSpace(dataWorksheet.Cells[currRow, 1]?.Value?.ToString() ?? string.Empty))
                    {
                        currValue = dataWorksheet.Cells[currRow, 1].Value.ToString();
                        if (currValue == "时段最大值")
                        {
                            maxValueRow = currRow;
                        }
                        else if (currValue == "时段最小值")
                        {
                            minValueRow = currRow;
                        }
                        else if (currValue == "时段平均值")
                        {
                            averageValueRow = currRow;
                        }
                        currRow++;
                    }

                    currCol = 2;    //第1列是"采集时间"
                    decimal tempDecimal;
                    while (!string.IsNullOrWhiteSpace(dataWorksheet.Cells[1, currCol]?.Value?.ToString() ?? string.Empty))
                    {
                        currChannel = dataWorksheet.Cells[1, currCol].Value.ToString();
                        dataList.Add(new MonitorData
                        {
                            Name = currChannel
                            ,
                            MaxValue = decimal.TryParse(dataWorksheet.Cells[maxValueRow, currCol]?.Value.ToString() ?? string.Empty, out tempDecimal) == true ? tempDecimal : errorValue
                            ,
                            MinValue = decimal.TryParse(dataWorksheet.Cells[minValueRow, currCol]?.Value.ToString() ?? string.Empty, out tempDecimal) == true ? tempDecimal : errorValue
                            ,
                            AverageValue = decimal.TryParse(dataWorksheet.Cells[averageValueRow, currCol]?.Value.ToString() ?? string.Empty, out tempDecimal) == true ? tempDecimal : errorValue
                        });
                        currCol++;
                    }
                    for (int i = 0; i < inputList.Count; i++)
                    {
                        inputList[i].MaxValue = dataList.Where(x => x.Name == inputList[i].Name).FirstOrDefault().MaxValue;
                        inputList[i].MinValue = dataList.Where(x => x.Name == inputList[i].Name).FirstOrDefault().MinValue;
                        inputList[i].AverageValue = dataList.Where(x => x.Name == inputList[i].Name).FirstOrDefault().AverageValue;
                        worksheet.Cells[2, inputList[i].ColIndex].Value = inputList[i].MaxValue - inputList[i].BaseValue;
                        worksheet.Cells[3, inputList[i].ColIndex].Value = inputList[i].MinValue - inputList[i].BaseValue;
                        worksheet.Cells[4, inputList[i].ColIndex].Value = inputList[i].AverageValue - inputList[i].BaseValue;
                    }
                    excelPackage.Save();
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.Message);
                    Console.WriteLine($"数据处理未成功!错误信息{ex.Message}");
                }

                Console.WriteLine("数据处理完成!");

            }
            catch (Exception)
            {
                Console.WriteLine("数据处理失败!");
            }
            
        }
    }
}
