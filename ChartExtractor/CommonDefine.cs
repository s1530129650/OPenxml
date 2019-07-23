using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace ChartExtractor
{
    class CommonDefine
    {
        public static void DumpJson(string fileName, object o, JsonSerializer serializer)
        {
            using (JsonWriter writer = new JsonTextWriter(new StreamWriter(fileName, false, Encoding.GetEncoding("UTF-8"))))
            {
                serializer.Serialize(writer, o);
            }
        }
    }
    public class DataSerializer
    {
        public static JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            ContractResolver = new DefaultContractResolver
            {
                NamingStrategy = new CamelCaseNamingStrategy()
            },
            NullValueHandling = NullValueHandling.Ignore,
            MissingMemberHandling = MissingMemberHandling.Ignore,
            Converters = new List<JsonConverter> { new CUnit.Converter()
            }
        };
        public static JsonSerializer Instance = JsonSerializer.CreateDefault(Settings);
    }
   

    [Serializable]
    public class FileInfor
    {
        public string Uid { get; set; }
        public FileHashStrings Hash { get; set; }
        public IList<string> CUids { get; set; }


    }
    public class FileHashStrings
    {
        public string Md5 { get; set; }
        public string Sha1 { get; set; }
        public string Sha256 { get; set; }

        public static FileHashStrings CalculateHashStrings(string filePathName)
        {
            var bytes = File.ReadAllBytes(filePathName);
            return new FileHashStrings()
            {
                Md5 = Convert.ToBase64String(System.Security.Cryptography.MD5.Create().ComputeHash(bytes)),
                Sha1 = Convert.ToBase64String(System.Security.Cryptography.SHA1.Create().ComputeHash(bytes)),
                Sha256 = Convert.ToBase64String(System.Security.Cryptography.SHA256.Create().ComputeHash(bytes))
            };
        }

    }
    //chart.xml -> c:chartspace -> c:chart
    [Serializable]
    public class CChart:IEquatable<CChart>
    {
        public string CUid { get; set; }
        public string Title { get; set; }
        public string Type { get; set; }
        public ChPlotArea PlotArea { get; set; }

        public static CChart GetInstance(A.Charts.ChartSpace chartSpace)
        {
            CChart cch = new CChart();
            // c:chart
            A.Charts.Chart chart = chartSpace.Descendants<A.Charts.Chart>().FirstOrDefault();
            //chart type
            string chrtype = chart?.PlotArea?.ChildElements[1]?.LocalName;
            cch.Type = chrtype;
            //chart title
            cch.Title = chart?.Title?.InnerText;
            //c:PlotArea
            cch.PlotArea = ChPlotArea.GetInstance(chartSpace);
            return  cch;
          
        }

        bool IEquatable<CChart>.Equals(CChart other)
        {
            throw new NotImplementedException();
        }
    }
    //chart.xml -> c:chartspace -> c:chart -> c:plotArea
    [Serializable]
    public class ChPlotArea:IEquatable<ChPlotArea>
    {  
        public string CatAxis { get; set; }
        public string ValAxis { get; set; }
        // c:plotArea=> c:XXchart => c:ser
        public IList<CSeries> Series { get; set; }
        public static ChPlotArea GetInstance(A.Charts.ChartSpace chartSpace)
        {
            ChPlotArea chplotArea = new ChPlotArea();
            var cSeries = new List<CSeries>();
            //c:chart->c:plotArea->c:catAx
            A.Charts.CategoryAxis catAxis = chartSpace.Descendants<A.Charts.CategoryAxis>().FirstOrDefault();
            chplotArea.CatAxis = catAxis?.Title?.InnerText;       
            //c:chart->plotArea->c:valAx
            A.Charts.ValueAxis valAxis = chartSpace.Descendants<A.Charts.ValueAxis>().FirstOrDefault();
            chplotArea.ValAxis = valAxis?.Title?.InnerText;
            //c:ser
           
            List<A.Charts.SeriesText>  serTxts = chartSpace.Descendants<A.Charts.SeriesText>().ToList() ;
            int cnt = serTxts.Count;
            if (cnt == 0)
            {
                A.Charts.SeriesText serTxt = chartSpace.Descendants<A.Charts.SeriesText>().FirstOrDefault();
                cSeries.Add(CSeries.GetInstance(chartSpace, serTxt));
            }
            else {
                foreach (var serTxt in serTxts)
                {
                    cSeries.Add(CSeries.GetInstance(chartSpace, serTxt));
                }
            }
            chplotArea.Series = cSeries;
            return chplotArea;
        }
        public bool Equals(ChPlotArea other)
        {
            throw new NotImplementedException();
        }
    }
    // c:plotArea=> c:XXchart => c:ser
    [Serializable]
    public class CSeries: IEquatable<CSeries>
    {
        public string Txt { get; set; }
        public string TxtRef { get; set; }
        public CCat Cat;
        public CVal Val;
        public static CSeries GetInstance(A.Charts.ChartSpace chartSpace, A.Charts.SeriesText serTxt)
        {
            CSeries cseries = new CSeries
            {
                //  c:ser -> c:tx ->c:v 
                Txt = serTxt?.InnerText
            };
            // c:ser -> c:tx -> c:strRef
            if (serTxt != null)
            {
                if (serTxt.InnerXml.Contains("<c:strRef"))
                {
                    cseries.Txt = serTxt.StringReference.Formula.InnerText;
                }
                else
                {
                    cseries.Txt = null;
                }
            }
            cseries.Cat = CCat.GetInstance(chartSpace);
            cseries.Val = CVal.GetInstance(chartSpace);

            return cseries;

        }
        public bool Equals(CSeries other)
        {
            throw new NotImplementedException();
        }
    }
    [Serializable]
    public class CCat:IEquatable<CCat>
    {
        public string CatStrRef { set; get; }
        public IList<CStrCache> StrCache;
        public static CCat GetInstance(A.Charts.ChartSpace chartSpace)
        {
            return null;
        }
            public bool Equals(CCat other)
        {
            throw new NotImplementedException();
        }
    }
    [Serializable]
    public class CStrCache:IEquatable<CStrCache>
    {
        public string PTCount { get; set; }
        public IList<CUnit> Unit;

        public bool Equals(CStrCache other)
        {
            throw new NotImplementedException();
        }
    }
    [Serializable]
    public class CUnit : IEquatable<CUnit>
    {
        public string Idx { get; set; }
        public string V { get; set; }
        public class Converter : JsonConverter<CUnit>
        {
            public override CUnit ReadJson(JsonReader reader, Type objectType, CUnit existingValue, bool hasExistingValue, JsonSerializer serializer)
            {
                List<string> pair = JToken.Load(reader).ToObject<List<string>>(serializer);
                return new CUnit()
                {
                    Idx = pair[0],
                    V   = pair[1]
                };
            }

            public override void WriteJson(JsonWriter writer, CUnit value, JsonSerializer serializer)
            {
                writer.WriteStartArray();
                writer.WriteValue(value.Idx);
                writer.WriteValue(value.V);
                writer.WriteEndArray();
            }
        }
        public bool Equals(CUnit other)
        {
            throw new NotImplementedException();
        }
    }
    [Serializable]
    public class CVal : IEquatable<CVal>
    {
        public string NumReference { set; get; }
        public CNumCache NumCache;
        public static CVal GetInstance(A.Charts.ChartSpace chartSpace)
        {
            CVal cval = new CVal();
            
            // c:chart->plotArea->c:type of chart(lineChart etc)-> c:ser -> c:numRef
            A.Charts.Values vals = chartSpace.Descendants<A.Charts.Values>().FirstOrDefault();
            cval.NumReference = vals?.NumberReference?.Formula?.InnerText;
           
            cval.NumCache = CNumCache.GetInstance(vals);
            return cval;

        }
        public bool Equals(CVal other)
        {
            throw new NotImplementedException();
        }
    }
    [Serializable]
    public class CNumCache:IEquatable<CNumCache>
    {
        public string FormatCode { get; set; }
        public string PTCount { get; set; }
        public IList<CUnit> Unit;
        public static CNumCache GetInstance(A.Charts.Values vals)
        {
            CNumCache cNumCache = new CNumCache();
            var chUnit = new List<CUnit>();
            //  c:ser -> c:numRef -> c:f
            cNumCache.FormatCode = vals?.NumberReference?.NumberingCache?.FormatCode?.InnerText;
            //  c:ser -> c:numRef -> c:numbercache —> c:pt
            cNumCache.PTCount = vals?.NumberReference?.NumberingCache?.PointCount?.Val;

            int count = 0;
            List<A.Charts.NumericPoint> numPt = vals?.NumberReference?.NumberingCache?.Descendants<A.Charts.NumericPoint>()?.ToList();
            foreach (var temp in numPt)
            {
                CUnit cunit = new CUnit
                {
                    Idx = count.ToString(),
                    V = temp.NumericValue.Text
                };
                count++;
                chUnit.Add(cunit);
            }
           
            
            cNumCache.Unit = chUnit;
            return cNumCache;

        }
            public bool Equals(CNumCache other)
        {
            throw new NotImplementedException();
        }
    }
}
