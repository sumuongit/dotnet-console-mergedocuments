using System.IO.Compression;
using System.Xml.Linq;

namespace MergeDocuments.Services
{
    public class MergeDocument
    {
        private readonly DynamicFooter _dynamicFooter;

        public MergeDocument(DynamicFooter dynamicFooter)
        {
            _dynamicFooter = dynamicFooter;
        }

        // Merges multiple DOCX files into one with dynamic footers for each appended document
        public void MergeDocsWithDynamicFooters(string[] files, string outputFile)
        {
            if (files == null || files.Length == 0) 
                throw new ArgumentException("No files provided.");

            foreach (var file in files)
            {
                if (!File.Exists(file))
                    throw new FileNotFoundException($"Input file not found: {file}");
            }

            if (string.IsNullOrWhiteSpace(outputFile))
                throw new ArgumentException("Output file path cannot be null or empty.", nameof(outputFile));

            // Copy first document as base for merge
            File.Copy(files[0], outputFile, overwrite: true);

            using var fs = new FileStream(outputFile, FileMode.Open, FileAccess.ReadWrite);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Update);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            // Load main document parts and XML
            var mainDocEntry = zip.GetEntry("word/document.xml") ?? throw new Exception("Main document not found.");
            var mainDoc = LoadXml(mainDocEntry);
            var body = mainDoc.Root.Element(w + "body");

            // Remove existing top-level section properties (will add per section later)
            body.Element(w + "sectPr")?.Remove();

            var relsDoc = LoadXml(zip.GetEntry("word/_rels/document.xml.rels") ?? throw new Exception("Rels file not found."));
            var contentTypesDoc = LoadXml(zip.GetEntry("[Content_Types].xml") ?? throw new Exception("Content types file not found."));

            // Clean up any existing footer relationships and parts
            RemoveExistingFooters(zip, relsDoc, contentTypesDoc);

            int footerCounter = 0;

            // Append content from each source DOCX, adding a dynamic footer per document
            for (int i = 0; i < files.Length; i++)
            {
                using var srcFs = new FileStream(files[i], FileMode.Open, FileAccess.Read);
                using var srcZip = new ZipArchive(srcFs, ZipArchiveMode.Read);

                var srcDoc = LoadXml(srcZip.GetEntry("word/document.xml") ?? throw new Exception("Source doc missing."));
                var srcBody = srcDoc.Root.Element(w + "body");
                var nodesToCopy = srcBody.Elements().Where(e => e.Name != w + "sectPr").Select(e => new XElement(e)).ToList();

                // Replace content if first file, otherwise append
                if (i == 0)
                    body.ReplaceAll(nodesToCopy);
                else
                    body.Add(nodesToCopy);

                footerCounter++;
                string relId = $"rIdFooter{footerCounter}";
                string footerFileName = $"footer{footerCounter}.xml";

                // Create footer part and update relationships/content types
                _dynamicFooter.CreateFooter(zip, relsDoc, contentTypesDoc, files[i], relId, footerFileName);

                // Add section break with footer reference between appended documents (except after last)
                if (i < files.Length - 1)
                {
                    body.Add(new XElement(w + "p", new XElement(w + "pPr",
                        new XElement(w + "sectPr",
                            new XElement(w + "footerReference",
                                new XAttribute(w + "type", "default"),
                                new XAttribute(XNamespace.Get(r.NamespaceName) + "id", relId))))));
                }
            }

            // Add final section properties with footer for the last document
            string finalFooterId = $"rIdFooter{footerCounter}";
            body.Add(new XElement(w + "sectPr",
                new XElement(w + "footerReference",
                    new XAttribute(w + "type", "default"),
                    new XAttribute(XNamespace.Get(r.NamespaceName) + "id", finalFooterId))));

            // Save updated rels, content types, and main document XML back to archive
            SaveXml(zip, "word/_rels/document.xml.rels", relsDoc);
            SaveXml(zip, "[Content_Types].xml", contentTypesDoc);
            SaveXml(zip, "word/document.xml", mainDoc);
        }

        // Load XML document from a ZipArchiveEntry
        private static XDocument LoadXml(ZipArchiveEntry entry)
        {
            using var reader = new StreamReader(entry.Open());
            return XDocument.Load(reader);
        }

        // Save XML document back into the ZipArchive at specified path
        private static void SaveXml(ZipArchive zip, string path, XDocument doc)
        {
            zip.GetEntry(path)?.Delete();
            var newEntry = zip.CreateEntry(path);
            using var writer = new StreamWriter(newEntry.Open());
            doc.Save(writer);
        }

        // Remove all existing footer relationships and parts from the ZIP archive
        private static void RemoveExistingFooters(ZipArchive zip, XDocument relsDoc, XDocument contentTypesDoc)
        {
            var footerRels = relsDoc.Root.Elements()
                .Where(e => (string)e.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
                .ToList();

            foreach (var rel in footerRels)
            {
                var target = (string)rel.Attribute("Target");
                rel.Remove();
                zip.GetEntry("word/" + target)?.Delete();

                var overrideElem = contentTypesDoc.Root.Elements().FirstOrDefault(e =>
                    e.Name.LocalName == "Override" && (string)e.Attribute("PartName") == "/word/" + target);
                overrideElem?.Remove();
            }
        }
    }
}