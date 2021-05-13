using System;
using System.IO;

namespace exeltoxml
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Старт конвертации");
            string importFilePath = @"C:\Users\iterator_pro\Desktop\allo_price.xlsx";

            ExelData exelData = new ExelData(importFilePath);
            string exportFilePath = @"C:\Users\iterator_pro\Desktop\allo_price.xml";
            string time = "2021-05-12 17:05";
            //string head = "This XML file does not appear to have any style information associated with it. The document tree is shown below.\n" +
            string head = "<?xml version = \"1.0\" encoding = \"UTF-8\" ?>\n" +
                      //    "<!DOCTYPE yml_catalog SYSTEM \"shops.dtd\" >\n" +
                          "<yml_catalog date = \"" + time + "\" >\n" +
                          "<shop>\n" +
                          "<name>ЕСТ</name>\n" +
                          "<company>ЕСТ</company>\n" +
                          "<currencies>\n" +
                          "<currency id = \"UAH\" rate = \"1\"/>\n" +
                          "</currencies>\n";

            string category = "";

            //create categories from exel file

            string categoryHead = "<categories>\n";
            string categoryFooter = "\n</categories>\n";
            string categories = categoryHead + category + categoryFooter;

            string offer = "";

            //create offers from exel file

            string offerHead = "<offers>\n";
            string offerFootter = "\n</offers>\n";
            string offers = offerHead + offer + offerFootter;

            string body = categories + offers;
            string footer = "</shop>\n" +
                            "</yml_catalog>";

            string text = head + body + footer;

            using (FileStream fstream = new FileStream(exportFilePath, FileMode.OpenOrCreate))
            {
                // преобразуем строку в байты
                byte[] array = System.Text.Encoding.Default.GetBytes(text);
                // запись массива байтов в файл
                fstream.Write(array, 0, array.Length);
                Console.WriteLine("Текст записан в файл " + exportFilePath);
            }
        }
    }
}
