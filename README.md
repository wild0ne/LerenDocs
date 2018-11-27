# What is it?
Leren is a simple Excel-based reporting engine. It helps to make a report in excel in fast-and-easy-to-make manner. You have to describe data sources right inside a cell of excel worksheet. It supports multi-level data structures and poor formatting options.

# Fast start!
1. Make an excel file, e.g. "file1.xlsx"
2. Type in cell A1 of excel: `{COLL=Apples/Green;WIDTH=2}`
3. Type in cell A2 of excel: `{Size}`
4. Let's write down data model:
```
public class GreenApple
{
   public double Size {set;get;}
}
public class Apples
{
  public List<GreenApple> Green {set;get;}
}
public class Root
{
  public Apples Apples {set;get;}
}
```
5. Ok, it's time to generate a report
```
// add some data
var root = new Root();
root.Apples = new Apples();
root.Apples.Green = new List<GreenApple>();
root.Apples.Green.Add(new GreenApple{Size = 5});
root.Apples.Green.Add(new GreenApple{Size = 10});

// making report
IReportEngine engine = new Engine();
engine.Provider = new ReflectionProvider(root);
engine.Go(@"file.xlsx", @"report.xlsx");
```
Now open the file `report.xlsx` to see what's generated.

# How it works?

## Data provder
At first, select a data provider. Some of them are ready out-of-box:
- Reflection provider for data stored in a tree of objects
- Xml Provider for data stored in xml file
- MySql provider for data stored in MySql database
- Oracle provider for data stored in Oracle database

## Custom data provider
If you aren't satisfied with capabailities that are ready out-of-the-box you can always implement data provider by yourself. Implement **IProvider** interface in order to use your own data provder.

```
public interface IProvider
    {
        int GetCollectionCount(string path, string tag, List<ContextItem> context);
        object GetValue(string path, string tag, List<ContextItem> context);
        ImageInfo GetImage(string path, string tag, List<ContextItem> context);
    }
```
##Report definition language
{COLL=Root/SomeProperty/SomeCollection}

**TODO: write more docs!**
