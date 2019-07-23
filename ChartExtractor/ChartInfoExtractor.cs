using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ChartExtractor
{
    public class ChartInfoExtractor
    {

        public static FileInfor Extract(string filePath, string uid)
        {
            FileInfor info = new FileInfor()
            {
                Uid = uid,
                Hash = FileHashStrings.CalculateHashStrings(filePath),
                CUids = new List<string>(),

            };
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                // xl\styles.xml numFmtId:formatCode <numFmt formatCode="_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"??_);_(@_)" numFmtId="164"/>
                Dictionary<uint, string> formatMappings = ExtractNumberFormatMappings(spreadSheetDocument.WorkbookPart);
                HashSet<CChart> cChart = new HashSet<CChart>();

                int chartCnt = 0;
                int incompleteChartNumber = 0;
                //int discardedChartNumber = 0;

                foreach (WorksheetPart worksheetPart in spreadSheetDocument.WorkbookPart.WorksheetParts)
                { 
                    // Deal with each chart
                    foreach (ChartPart chartPart in worksheetPart.DrawingsPart.ChartParts)
                    {
                        if (chartPart == null || chartPart.ChartSpace == null)
                        {
                            incompleteChartNumber++;
                            continue;
                        }
                        // whether this chart is discarded
                        //bool isDiscarded = false;
                        //chart1.xml  c:chartSpace
                        A.Charts.ChartSpace chartSpace = chartPart.ChartSpace;
                        
                        if (chartSpace.Count() < 1)
                        {
                            incompleteChartNumber++;
                            continue;
                        }
                        //c:chartSpace -> c:chart
                        
                        var newChart = CChart.GetInstance(chartSpace);
                        string cUid = uid + $".ch{chartCnt}";
                        info.CUids.Add(cUid);
                        newChart.CUid = cUid;
                        CommonDefine.DumpJson(newChart.CUid + ".json", newChart, DataSerializer.Instance);
                        chartCnt++;

                    }

                }

            }

            CommonDefine.DumpJson(uid + ".json", info, DataSerializer.Instance);
            return info;

        }

        private static Dictionary<uint, string> ExtractNumberFormatMappings(WorkbookPart workbookPart)
        {
            Dictionary<uint, string> formatMappings = new Dictionary<uint, string>();

            var numFormatsParentNodes = workbookPart.WorkbookStylesPart.Stylesheet.ChildElements.OfType<NumberingFormats>();
            foreach (var parentNode in numFormatsParentNodes)
            {
                var formatNodes = parentNode.ChildElements.OfType<NumberingFormat>();
                foreach (var formatNode in formatNodes)
                {
                    uint id = formatNode.NumberFormatId.Value;
                    if (formatMappings.ContainsKey(id) && formatNode.FormatCode.Value != formatMappings[id])
                        throw new ArgumentException("An item with the same key but different value has already been added.");
                    formatMappings[id] = formatNode.FormatCode;
                }
            }
            return formatMappings;
        }



    }
}
