using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace readChart
{
    class Program
    {
        
        
        public void GetChart(string fileName)
        {
            string txt;
            
            //string fileName = @"D:\c#file\excelfile\test1.xlsx";
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                //Console.WriteLine("*********"+ doc.WorkbookPart);
                WorkbookPart bkPart = doc.WorkbookPart;
                Workbook workbook = bkPart.Workbook;
                //Console.WriteLine(workbook.Ancestors());
                
                //Sheet s = workbook.Descendants<Sheet>().Where(sht => sht.Name == "Sheet1").FirstOrDefault();
                //Sheet s = workbook.Descendants<Sheet>().Where(sht => sht.Name == "Razem").FirstOrDefault();
                //IEnumerable s1 = workbook.Descendants<Sheet>();
                //Console.WriteLine(workbook.Descendants<Sheet>());
                Sheet s = workbook.Descendants<Sheet>().FirstOrDefault();
                //Console.WriteLine(s);
                //Console.WriteLine(s.Id);
                //Console.WriteLine("$$$$$$$");
                WorksheetPart wsPart = (WorksheetPart)bkPart.GetPartById(s.Id);  
                DrawingsPart dp = (DrawingsPart)wsPart.DrawingsPart;
                /*
                Console.WriteLine("$$$$$$$");
                Console.WriteLine(dp.ChartParts);
                Console.WriteLine("$$$$$$$");
                foreach (var chartPart in dp.ChartParts) {
                    Console.WriteLine(chartPart.ChartColorStyleParts);
                    Console.WriteLine("$$$$$$$");
                }*/


                WorksheetDrawing dWs = dp.WorksheetDrawing;

               
                Console.WriteLine("The count of the charts is : "+dWs.ChildElements.Count);
                //Console.WriteLine(dWs.ChildElements[1]);
                //Console.WriteLine(dWs);
                Console.WriteLine("**********************");
                //Console.WriteLine(dWs.Descendants());

               
                txt = dWs.Descendants<A.Spreadsheet.NonVisualDrawingProperties>().FirstOrDefault().Name;
                Console.WriteLine("The name(not title) of the charts is : " + txt);              
                Console.WriteLine(dWs.Descendants<A.Spreadsheet.NonVisualDrawingProperties>().Count());
                //Console.WriteLine(dWs.Descendants<A.Spreadsheet.NonVisualDrawingProperties>().FirstOrDefault().ChildElements + "############");
                txt = dWs.Descendants<A.Spreadsheet.NonVisualDrawingProperties>().ElementAtOrDefault(0).Name;
                //Console.WriteLine(txt );
                Console.WriteLine("**********************");
          
                //Console.WriteLine(dp.ChartParts.Count());
                ChartPart cp = dp.ChartParts.FirstOrDefault();
                Console.WriteLine("the chart space language is :" + cp.ChartSpace.EditingLanguage.Val);
                Console.WriteLine("**********************");
                //Console.WriteLine("the XXXXX is :" + cp.ChartSpace.RoundedCorners);

                A.Charts.ChartShapeProperties cs = cp.ChartSpace.Descendants<A.Charts.ChartShapeProperties>().FirstOrDefault();
                //Console.WriteLine("the property of the Charts :" );
                //Console.WriteLine(cs.LocalName);
              
                A.Charts.AxisDataSourceType adst = cp.ChartSpace.Descendants<A.Charts.AxisDataSourceType>().FirstOrDefault();
                Console.Write("the reference of the catagory is :");
                //Console.WriteLine(adst.StringReference.Formula.InnerText);  // if there is a reference 
                Console.WriteLine("the reference of the  Number is :");
                //Console.WriteLine(adst.NumberReference.Formula.InnerText);  // if there is a reference 
                Console.WriteLine("**********************");

                A.Charts.CategoryAxis ca = cp.ChartSpace.Descendants<A.Charts.CategoryAxis>().FirstOrDefault();
                Console.WriteLine("the title of the Category axix is :");
                Console.WriteLine(ca.Title.InnerText);

                A.Charts.ValueAxis va = cp.ChartSpace.Descendants<A.Charts.ValueAxis>().FirstOrDefault();
                Console.WriteLine("the title of the Value axix is :");
                Console.WriteLine(va.Title.InnerText);

                A.Charts.Chart c = (A.Charts.Chart)cp.ChartSpace.Descendants<A.Charts.Chart>().FirstOrDefault();
                //Console.WriteLine(c.LocalName);

               


                Console.WriteLine("**********************");
                Console.WriteLine("the title of the chart is :" + c.Title.InnerText);

                //Console.WriteLine( c.LocalName);
                //Console.WriteLine("chart title is :" + c.Title.InnerXml);
                //Console.WriteLine("chart title is :" + c.Title.ChartText.RichText.InnerText);
                Console.WriteLine("**********************");
                Console.WriteLine("the type of the chart is :" + c.PlotArea.ChildElements[1].LocalName.ToString());
                Console.WriteLine("**********************");
                Console.WriteLine("the other of the chart is :" + c.PlotArea.LocalName);
              
                //Console.WriteLine("the type of the chart is :" + c.PlotArea.ChildElements[1].ChildElements.Count);


                //Console.WriteLine(txt + "############");
                //Console.WriteLine(c.PlotArea.ChildElements[2].LocalName.ToString() + "############");
                //Console.WriteLine(c.PlotArea.ChildElements[3].LocalName.ToString());
                //Console.WriteLine(c.PlotArea.ChildElements[4].LocalName.ToString());
                //Console.WriteLine(c.PlotArea.ChildElements[5].LocalName.ToString());
                Console.ReadKey();
                //this.txtMessages.Text = txt;
            }


        }
        
        static void Main(string[] args)
        {
            Program program = new Program();
            //program.GetChart(@"D:\common_analysis\downloadfile\second.XLSX");
           program.GetChart(@"D:\common_analysis\testfile\test2nd.xlsx");
            //program.GetChart(@"D:\common_analysis\testfile\new.xlsx");
        }
    }
}
