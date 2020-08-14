using Leren.Lerengine;
using Leren.Providers.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Formulas
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
                engine.Go(getResourceData(@"Formulas.template.xlsx"), fs);
            }
            Process.Start(path);
        }

        static object GetModel()
        {
            const int rowsCount = 10;

            var result = new List<object>();
            for (int i = 1; i <= rowsCount; i++)
            {
                result.Add(new { Vendor = "Tesla", Model = $"Cubertruck {i}", Price = i * 10, Number = i });
            }

            return new
            {
                Data = result,
                Header = $"This sample demonstrates the way of using excel built-in formulas",
                Header2 = $"SUM and AVG are used"
            };
        }

        static private Stream getResourceData(string resourceName)
        {
            var assembly = Assembly.Load("Formulas");
            var names = assembly.GetManifestResourceNames();

            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
