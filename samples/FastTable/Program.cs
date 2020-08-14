using Leren.Lerengine;
using Leren.Providers.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace FastTable
{
    class Program
    {
        static void Main(string[] args)
        {
            var engine = new Engine();
            engine.Provider = new ReflectionProvider(GetModel());
            string path = Path.GetTempFileName() + ".xlsx";
            using (FileStream fs = new FileStream(path, FileMode.CreateNew))
            {
                engine.OnProgress += (s, pea) => Console.Title = string.Format("{0:000.0} %", pea.ProgressPercent);
                engine.Go(getResourceData(@"FastTable.template.xlsx"), fs);
            }
            Process.Start(path);
        }

        static object GetModel()
        {
            const int rowsCount = 100 * 1000;

            var result = new List<object>();
            for (int i = 1; i <= rowsCount; i++)
            {
                result.Add(new { Vendor = "Tesla", Model = $"Cubertruck {i}", Number = i });
            }

            return new { Data = result, Header = $"This sample demonstrates the way to achieve maximum performance. Note that report reading and editing stays easy. Rows count is {rowsCount}",
                Header2 = $"Excel report generation started at {DateTime.Now}, finished now" };
        }

        static private Stream getResourceData(string resourceName)
        {
            var assembly = Assembly.Load("FastTable");
            var names = assembly.GetManifestResourceNames();

            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
