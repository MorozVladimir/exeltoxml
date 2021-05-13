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
        //  public List<UserForMailing> users { get; set; }
        //    public List<Colums12> datas { get; set; }
        //    List<ForWeight> forWeights = new List<ForWeight>();


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

            }











            //  ProgressBar button = (ProgressBar)d;
            //  button.Value = 30.0;
            //users = new List<UserForMailing>();
            //  datas = new List<Colums12>();
            //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            //// System.Text.Encoding.Convert.;
            //using (var stream = File.Open(fName, FileMode.Open, FileAccess.Read))
            //{
            //    using (var reader = ExcelReaderFactory.CreateReader(stream))
            //    {
            //        int count = 0;
            //        do
            //        {
            //            //   var ufm = new UserForMailing();
            //            //  var data = new Colums12();
            //            while (reader.Read()) //Each ROW
            //            {
            //                count++;
            //                string firstname = "";
            //                for (int column = 0; column < reader.FieldCount; column++)
            //                {
            //                    //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
            //                    //    Console.WriteLine(reader.GetValue(column));//Get Value returns object
            //                    // string rowIs;
            //                    switch (column)
            //                    {
            //                        case 0:
            //                            var di = reader.GetValue(column).ToString();
            //                            //    data.Column0 = di;
            //                            //switch (d)
            //                            //{
            //                            //    case "0":
            //                            //        rowIs = "category";
            //                            //        break;
            //                            //    case "1":
            //                            //        rowIs = "good";
            //                            //        break;
            //                            //    case "2":

            //                            //        break;
            //                            //    case "3":
            //                            //        rowIs = "groupCategory";
            //                            //        break;
            //                            //    default:

            //                            //        break;
            //                            //}
            //                            //  ufm.Barcode = reader.GetValue(column).ToString();
            //                            break;
            //                        case 1:
            //                            try
            //                            {
            //                                //       data.Column1 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 2:
            //                            try
            //                            {
            //                                //       data.Column2 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 3:
            //                            try
            //                            {
            //                                //       data.Column3 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 4:
            //                            try
            //                            {
            //                                //        data.Column4 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 5:
            //                            try
            //                            {
            //                                //        data.Column5 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 6:
            //                            try
            //                            {
            //                                var dgh = reader.GetValue(column).ToString();
            //                                var hh = dgh.Split(",");
            //                                //       data.Column6 = hh[0] + "." + hh[1];
            //                                // data.Column6 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 7:
            //                            try
            //                            {
            //                                //       data.Column7 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            //  firstname = reader.GetValue(column).ToString();
            //                            break;
            //                        case 8:
            //                            try
            //                            {
            //                                //       data.Column8 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            //  ufm.Name = firstname + " " + reader.GetValue(column).ToString();
            //                            break;
            //                        case 9:
            //                            try
            //                            {
            //                                //         data.Column9 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 10:
            //                            try
            //                            {
            //                                //        data.Column10 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                        case 11:
            //                            try
            //                            {
            //                                //        data.Column11 = reader.GetValue(column).ToString();
            //                            }
            //                            catch
            //                            {

            //                            }
            //                            break;
            //                            //    case 13:
            //                            //  ufm.Email = reader.GetValue(column).ToString();
            //                            //       break;
            //                            //   case 14:
            //                            // ufm.ExhibitionName = reader.GetValue(column).ToString();
            //                            //      break;
            //                    }
            //                    //           button.Value = button.Value + 1;
            //                }
            //                var persentage = (count * 100) / reader.RowCount;
            //                //   progress.Report(persentage);
            //                // d.Value = count;
            //                //  datas.Add(data);
            //                //   data = new Colums12();
            //                // users.Add(ufm);
            //                // ufm = new UserForMailing();
            //            }
            //        }
            //        while (false); //Move to NEXT SHEET

            //    }
            //}
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