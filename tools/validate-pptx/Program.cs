using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

// Usage: validate-pptx <file.pptx> [--json]
// Prints one line per problem the Open XML SDK finds, the same class of problem
// that makes PowerPoint offer to "repair" a file.
if (args.Length < 1)
{
    Console.Error.WriteLine("usage: validate-pptx <file.pptx>");
    return 2;
}

try
{
    using var document = PresentationDocument.Open(args[0], false);
    var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
    var count = 0;
    foreach (var error in validator.Validate(document))
    {
        count++;
        var part = error.Part?.Uri.ToString() ?? "?";
        var path = error.Path?.XPath ?? "?";
        // Tabs separate the fields so callers can split reliably.
        Console.WriteLine($"{error.ErrorType}\t{part}\t{path}\t{error.Id}\t{error.Description.Replace('\n', ' ')}");
    }
    Console.Error.WriteLine($"{count} problem(s)");
    return 0;
}
catch (Exception exception)
{
    // A package the SDK cannot even open is the worst outcome of all.
    Console.WriteLine($"Fatal\t{args[0]}\t-\t{exception.GetType().Name}\t{exception.Message.Replace('\n', ' ')}");
    return 1;
}
