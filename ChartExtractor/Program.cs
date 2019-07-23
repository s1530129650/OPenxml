using System;
using CommandLine;

namespace ChartExtractor
{
    class Program
    {
        public class Options
        {
            [Option('i', "input", Required = true)]
            public string InputFile { get; set; }

            [Option('u', "uid", Required = true)]
            public string Uid { get; set; }
        }
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(o =>
            {

                Console.WriteLine($"Extracting PivotTable from file {o.InputFile} to {o.Uid}.");

                ChartInfoExtractor.Extract(o.InputFile, o.Uid);

            });
        }
    }
}
