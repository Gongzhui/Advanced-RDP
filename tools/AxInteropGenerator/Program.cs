using System;
using System.IO;
using System.Linq;
using System.Reflection;

if (args.Length < 1)
{
    Console.WriteLine("Usage: AxInteropGenerator <outputDirectory>");
    return 1;
}

var outputDir = Path.GetFullPath(args[0]);
Directory.CreateDirectory(outputDir);

var tlbPath = Path.Combine(Environment.SystemDirectory, "mstscax.dll");
if (!File.Exists(tlbPath))
{
    Console.Error.WriteLine($"mstscax.dll not found at {tlbPath}");
    return 2;
}

var designAssemblyPath = Directory.GetFiles(
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "shared", "Microsoft.WindowsDesktop.App"),
        "System.Windows.Forms.Design.dll",
        SearchOption.AllDirectories)
    .OrderByDescending(p => p)
    .FirstOrDefault();

if (designAssemblyPath == null)
{
    Console.Error.WriteLine("System.Windows.Forms.Design.dll not found in dotnet shared framework.");
    return 3;
}

var designAssembly = Assembly.LoadFrom(designAssemblyPath);
var axImporterType = designAssembly.GetType("System.Windows.Forms.Design.AxImporter")
    ?? throw new InvalidOperationException("AxImporter type not found.");
var optionsType = designAssembly.GetType("System.Windows.Forms.Design.AxImporter")?
    .GetNestedType("Options", BindingFlags.Public | BindingFlags.NonPublic)
    ?? throw new InvalidOperationException("AxImporter.Options type not found.");

var options = Activator.CreateInstance(optionsType)!;
optionsType.GetField("outputDirectory")?.SetValue(options, outputDir);
optionsType.GetField("genSources")?.SetValue(options, false);

var importer = Activator.CreateInstance(axImporterType, args: new[] { options })!;

var generateMethod = axImporterType.GetMethod("GenerateFromFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
if (generateMethod == null)
{
    Console.Error.WriteLine("GenerateFromFile method not found on AxImporter.");
    return 4;
}

generateMethod.Invoke(importer, new object[] { tlbPath });
Console.WriteLine($"Generated interop assemblies in {outputDir}");
return 0;
