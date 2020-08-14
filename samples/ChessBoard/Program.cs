using Leren.Lerengine;
using Leren.Providers.Reflection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ChessBoard
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
                engine.Go(getResourceData(@"ChessBoard.template.xlsx"), fs);
            }
            Process.Start(path);
        }

        static object GetModel()
        {
            return new
            {
                Rows = new List<object> {
                    new { Cols = new int[] { 0, 1, 0, 1, 0, 1, 0, 1 } },
                    new { Cols = new int[] { 1, 0, 1, 0, 1, 0, 1, 0 } },
                    new { Cols = new int[] { 0, 1, 0, 1, 0, 1, 0, 1 } },
                    new { Cols = new int[] { 1, 0, 1, 0, 1, 0, 1, 0 } },
                    new { Cols = new int[] { 0, 1, 0, 1, 0, 1, 0, 1 } },
                    new { Cols = new int[] { 1, 0, 1, 0, 1, 0, 1, 0 } },
                    new { Cols = new int[] { 0, 1, 0, 1, 0, 1, 0, 1 } },
                    new { Cols = new int[] { 1, 0, 1, 0, 1, 0, 1, 0 } },
                }
            };
        }

        static private Stream getResourceData(string resourceName)
        {
            var assembly = Assembly.Load("ChessBoard");
            var names = assembly.GetManifestResourceNames();

            return assembly.GetManifestResourceStream(resourceName);
        }
    }
}
