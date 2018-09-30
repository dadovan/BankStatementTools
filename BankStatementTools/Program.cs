using System;
using System.Linq;
using System.Reflection;

namespace BankStatementTools
{
    public static class Program
    {
        // TODO: Turn into tool
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                var toolName = $"{Assembly.GetEntryAssembly().GetName().Name}.exe";
                Console.WriteLine("Usage:");
                Console.WriteLine($"\tPS C:\\> {toolName} (InputPDF) (OutputCSV)");
                Console.WriteLine($"\tPS C:\\> {toolName} (InputPDF1) .. (InputPDFn)(OutputCSV)");
                Console.WriteLine();
                Console.WriteLine("Example:");
                Console.WriteLine($"\tPS C:\\> {toolName} C:\\Downloads\\JanStatement.pdf C:\\Documents\\JanStatement.csv");
                Console.WriteLine($"\tPS C:\\> {toolName} C:\\Downloads\\JanStatement.pdf C:\\Downloads\\FebStatement.pdf C:\\Downloads\\MarStatement.pdf C:\\Documents\\JanStatement.csv");
                return;
            }

            var inputs = args.Take(args.Length - 1).ToArray();
            var output = args.Last();
            SynovusStatementReader.TransformStatements(output, inputs);
        }
    }
}
