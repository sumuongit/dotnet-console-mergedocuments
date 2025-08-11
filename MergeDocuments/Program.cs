using MergeDocuments.Services;

class Program
{
    static void Main()
    {
        // Prepare input file paths relative to the app base directory
        string[] filesToMerge =
        {
           Path.Combine(AppContext.BaseDirectory, "Docs", "source1.docx"),
           Path.Combine(AppContext.BaseDirectory, "Docs", "source2.docx"),
        };

        // Output merged document path
        string outputPath = Path.Combine(AppContext.BaseDirectory, "Docs", "merged.docx");

        // Initialize merger service with dynamic footer functionality
        var merger = new MergeDocument(new DynamicFooter());
        merger.MergeDocsWithDynamicFooters(filesToMerge, outputPath);

        Console.WriteLine("Merged file created: " + outputPath);

        // Define content control replacements (tag -> replacement text)
        var replacements = new Dictionary<string, string>
        {
            { "ClientName", "ImpleVista" },
            { "AssignmentName", "MergeDocuments" }
        };

        // Initialize content control updater and apply replacements
        var contentUpdater = new UpdateContentControl();
        contentUpdater.UpdateContentControls(outputPath, replacements);
    }
}
