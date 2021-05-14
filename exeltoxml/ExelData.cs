using ExcelDataReader;
using exeltoxml.XMLConstruction;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;


namespace exeltoxml
{
    public class ExelData
    {
        public Company company { get; set; }
        public List<Category> categories { get; set; }
        public List<good> goods { get; set; }

        public ExelData(string fName)
        {
            company = new Company();

            // JArray data = new JArray();
            //       using (ExcelPackage package = new ExcelPackage(file.OpenReadStream()))
            FileInfo fileInfo = new FileInfo(fName);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet companysheet = package.Workbook.Worksheets[0];
                ExcelWorksheet categorysheet = package.Workbook.Worksheets[1];
                ExcelWorksheet goodsheet = package.Workbook.Worksheets[2];

                company.name = companysheet.Cells[1, 2].Value.ToString();
                company.url = companysheet.Cells[2, 2].Value.ToString();
                company.currenciesId = companysheet.Cells[3, 2].Value.ToString();
                company.currenciesRate = companysheet.Cells[4, 2].Value.ToString();
                //Process, read from excel here and populate jarray
                //  int companyRowCount = companysheet.Dimension.Rows;
                //   int companyColCount = companysheet.Dimension.Columns;

                Console.WriteLine(company.name);
                Console.WriteLine(company.url);
                Console.WriteLine(company.currenciesId);
                Console.WriteLine(company.currenciesRate);
                Console.WriteLine();


                categories = new List<Category>();
                int categoryRowCount = categorysheet.Dimension.Rows;
                int categoryColCount = categorysheet.Dimension.Columns;

                for(int i = 2; i <= categoryRowCount; i++)
                {
                    Category category = new Category();
                    category.id = Int32.Parse(categorysheet.Cells[i, 1].Value.ToString());
             //       category.parentId = Int32.Parse(categorysheet.Cells[i, 2].Value.ToString());
                    try { category.parentId = Int32.Parse(categorysheet.Cells[i, 2].Value.ToString()); }
                    catch { category.parentId = 0; }
                    category.value = categorysheet.Cells[i, 3].Value.ToString();
                    categories.Add(category);
                    
                }

                foreach(var cat in categories)
                {
                    Console.WriteLine(cat.id + " " + cat.parentId + " " + cat.value);
                }


                goods = new List<good>();
                int goodRowCount = goodsheet.Dimension.Rows;
                int goodColumnCount = goodsheet.Dimension.Columns;
                Console.WriteLine("grow -" + goodRowCount + "  gcol -" + goodColumnCount);
                bool nextId = false;
                good good = new good();
                for (int i = 2; i <= goodRowCount; i++)
                {
                    //       good = new good();
                    int gidtemp; 
                    try { gidtemp = Int32.Parse(goodsheet.Cells[i, 1].Value.ToString()); }
                    catch { gidtemp = 0; }
                    Console.Write("gid-" + gidtemp);
                    if (gidtemp != 0)
                    {
                        if (nextId == true)
                        {
                            goods.Add(good);
                            good = new good();
                            //                         nextId = true;
                        }
                        else
                        {
                            nextId = true;
                        }
                        good.id = gidtemp;
                        good.available = (goodsheet.Cells[i, 2].Value.ToString().Equals("true")) ? true : false;
                 //       Console.Write("  gava-" + good.available);
                        good.price = decimal.Parse(goodsheet.Cells[i, 3].Value.ToString());
                 //       Console.Write("  price-" + good.price);
                        good.priceOld = decimal.Parse(goodsheet.Cells[i, 4].Value.ToString());
                 //       Console.Write("  priceOld-" + good.priceOld);
                        good.pricePromo = decimal.Parse(goodsheet.Cells[i, 5].Value.ToString());
                 //       Console.Write("  pricePromo-" + good.pricePromo);
                        good.stockQuantity = decimal.Parse(goodsheet.Cells[i, 6].Value.ToString()); 
                //        Console.Write("  stockQuantity-" + good.stockQuantity);
                        good.CurrencyId = goodsheet.Cells[i, 7].Value.ToString();
               //         Console.Write("  currenciId-" + good.CurrencyId);
                        good.categoryId = Int32.Parse(goodsheet.Cells[i, 8].Value.ToString());
                //        Console.Write("  categoryId-" + good.categoryId);
                      //   good.pictures = new List<string>();
                      //  Console.Write("  picture-" + goodsheet.Cells[i, 9].Value.ToString());
                        good.pictures.Add(goodsheet.Cells[i, 9].Value.ToString());
                //        Console.Write("  picture-" + good.pictures.ToArray()[0]);
                        good.name = goodsheet.Cells[i, 10].Value.ToString();
                        good.article = goodsheet.Cells[i, 11].Value.ToString();
                        good.vendor = goodsheet.Cells[i, 12].Value.ToString();
                        good.description = goodsheet.Cells[i, 13].Value.ToString();
                        goodParam goodParam = new goodParam();
                        goodParam.name = goodsheet.Cells[i, 14].Value.ToString();
                        goodParam.value = goodsheet.Cells[i, 15].Value.ToString();
                        ///  good.pictures
                        Console.WriteLine(goodParam.name + " " + goodParam.value + " " + good.id);
                        good.parametrs.Add(goodParam);

                    }
                    else
                    {
                        try
                        {
                            var pic = goodsheet.Cells[i, 9].Value.ToString();
                            good.pictures.Add(pic);
                        }
                        catch
                        {

                        }

                        try
                        {
                            var parname = goodsheet.Cells[i, 14].Value.ToString();
                            var parval = goodsheet.Cells[i, 15].Value.ToString();
                        //    Console.WriteLine(parname + " " + parval);
                            goodParam goodParam = new goodParam();
                            goodParam.name = parname;
                            goodParam.value = parval;
                            Console.WriteLine(goodParam.name + " " + goodParam.value + " " + good.id);
                            good.parametrs.Add(goodParam);
                        }
                        catch
                        {

                        }
                        Console.WriteLine(i + " из " + goodRowCount);
                        if (i == goodRowCount)
                        {
                            goods.Add(good);
                        }
                    }
                }
                foreach(var g in goods)
                {
                    Console.WriteLine("gid-" + g.id + "  gname-" + g.name ); 
                    foreach(var pi in g.pictures)
                    {
                        Console.WriteLine("   ->" + pi);
                    }
                    foreach (var pa in g.parametrs)
                    {
                        Console.WriteLine("              paname>>" + pa.name + "   paval>>" + pa.value + " " + g.id);
                    }
                }

                Console.WriteLine("+++++++++++++++++++++");
                Console.WriteLine(goods.ToArray()[0].parametrs.ToArray()[0].value);
            }
        }

        private string validate1(string text)
        {
            var d0 = text.Split("\"");
            var res = d0[0] + "❜" + d0[1];
            if (res.Contains("\""))
            {
                validate1(res);
                return res;
            }
            else return res;
        }
        private string validate2(string text)
        {
            var d0 = text.Split("'");
            var res = d0[0] + "❜" + d0[1];
            if (res.Contains("'"))
            {
                validate2(res);
                return res;
            }
            else return res;
        }
        public void writeDataToExel(string filePathPLU, IProgress<int> progress, ExelData exelData)
        {
            //  ForWeight forWeight = new ForWeight();
            int count = 0;
            //foreach (var data in datas)
            //{
            //    count++;
            //    if (data.Column0.Equals("1"))
            //    {
            //        forWeight.plu = data.Column11;
            //        forWeight.namePlu = data.Column3;
            //        forWeight.namePLU2 = data.Column4;
            //        forWeight.articul = data.Column5;
            //        forWeight.group = Int32.Parse(data.Column2);
            //        forWeight.price = data.Column6;
            //        forWeight.expirationDate = data.Column7;
            //        forWeight.date = data.Column7;
            //        forWeight.pluType = "By Weight";
            //        forWeight.discount = "No";
            //        forWeight.freePrice = "No";
            //        forWeight.additionalText = "0";
            //        forWeight.buttonNumber = "0";
            //        forWeight.spesialProposition = "0.00";

            //        forWeights.Add(forWeight);
            //        forWeight = new ForWeight();
            //        progress.Report((count * 100) / datas.Count());
            //    }
            //}

            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

            //DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(forWeights), (typeof(DataTable)));
            //var memoryStream = new MemoryStream();

            //using (var fs = new FileStream(filePathPLU, FileMode.Create, FileAccess.Write))
            //{
            //    IWorkbook workbook = new XSSFWorkbook();
            //    ISheet excelSheet = workbook.CreateSheet("Sheet1");

            //    List<String> columns = new List<string>();
            //    IRow row = excelSheet.CreateRow(0);
            //    int columnIndex = 0;

            //foreach (System.Data.DataColumn column in table.Columns)
            //{
            //    columns.Add(column.ColumnName);
            //    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
            //    columnIndex++;
            //}

            //int rowIndex = 1;
            //foreach (DataRow dsrow in table.Rows)
            //{
            //    row = excelSheet.CreateRow(rowIndex);
            //    int cellIndex = 0;
            //    foreach (String col in columns)
            //    {
            //        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
            //        cellIndex++;
            //    }

            //    rowIndex++;
            //}
            //workbook.Write(fs);
        }


    }

}