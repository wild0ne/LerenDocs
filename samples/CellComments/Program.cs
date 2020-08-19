using Leren.Lerengine;
using Leren.Providers.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CellComments
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
                engine.Go(getResourceData(@"CellComments.template.xlsx"), fs);
            }
            Process.Start(path);
        }

        static object GetModel()
        {
            const int rowsCount = 10;

            var result = new List<object>();
            for (int i = 1; i <= rowsCount; i++)
            {
                result.Add(new { Value = i % 2 == 0 ? $"comment is here" : $"{i}: no comments", Note = i % 2 == 0 ? $"Row number is {i}" : null });
            }

            return new
            {
                Data = result,
            };
        }

        static private Stream getResourceData(string resourceName)
        {
            var assembly = Assembly.Load("CellComments");
            var names = assembly.GetManifestResourceNames();

            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
