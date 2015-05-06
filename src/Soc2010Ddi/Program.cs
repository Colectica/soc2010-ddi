using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Soc2010Ddi
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                ShowUsage();
                return 1;
            }

            string inputFileName = args[0];
            string outputFileName = args[1];

            if (!File.Exists(inputFileName))
            {
                Console.WriteLine("Sorry, the input file does not exist.");
                ShowUsage();
            }

            Soc2010ToDdiConverter converter = new Soc2010ToDdiConverter();
            converter.Convert(inputFileName, outputFileName);

#if DEBUG
            Console.WriteLine("Press enter to end the program.");
            Console.ReadLine();
#endif

            return 0;
        }

        static void ShowUsage()
        {
            Console.WriteLine("The SOC2010 to DDI tool describes the SOC classification using the DDI 3.2 XML standard.");
            Console.WriteLine("The tool reads the SOC classification from a spreadsheet with the following columns:");
            Console.WriteLine("  Major Group, Sub-Major Group, Minor Group, Unit Group, Group Title");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine();
            Console.WriteLine("  Soc2010ToDdi.exe inputFile outputFile");
        }
    }
}
