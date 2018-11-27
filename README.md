# What is it?
Leren is a simple Excel-based reporting engine. It helps to make a report in excel in fast-and-easy-to-make manner. You have to describe data sources right inside a cell of excel worksheet. It supports multi-level data structures and poor formatting options.

# Fast start!
1. Make an excel file, e.g. "file1.xlsx"
2. Type in cell A1 of excel: `{COLL=Apples/Green;WIDTH=2}`
3. Type in cell A2 of excel: `{Size}`
4. Let's write down data model:
```c#
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
```c#
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

```c#
public interface IProvider
    {
        int GetCollectionCount(string path, string tag, List<ContextItem> context);
        object GetValue(string path, string tag, List<ContextItem> context);
        ImageInfo GetImage(string path, string tag, List<ContextItem> context);
    }
```
## Report definition language
There is a syntax for data source definition: curvy bracets and special words. There is only two types of definitions: **collection** and **property**.

## Language for Reflection Provider

Use collection definition in order to make some cells repeat itself x times. Here is a sample of such definition:
```
{COLL=Root/SomeProperty/SomeCollection;HEIGHT=1;WIDTH=10;GROW=DOWN;NOINSERT='YES'}
```
Arguments are described below:
- COLL= is a path to property, each element (property) is separated with **"/"**, starting from the root of data model. When you place one collection inside another, you have to specify path to collection starting from current item (context).
- HEIGHT - ...
- WIDTH - ...
- GROW - grow direction. Use 'right' to make it grow right, or 'down' for growing down.
- NOINSERT - when it's set to 'yes', inserting of cells is not performed while processing current collection. Default value is 'no'. It is useful when you want to generate a chess board, for example.

Use property definition to display data. Here is full sample:
```
{Car1/Wheel1/Diam;MULT=3.1;ADD=100;FORMAT=0.000}
```
Arguments are described below:
- necessary argument is a path to property. Use fully-qualified path, starting from data root, or starting from current item, when you are in a collection context.
- MULT multiples property value by it's argument.
- ADD adds argument to property value or to result of MULT, if MULT is specified.
- FORMAT is formatting numeric property value.

## Language for Oracle/MySql Provider

bla-bla-bla

## Language for XML Provider

bla-bla-bla


**TODO: write more docs!**
