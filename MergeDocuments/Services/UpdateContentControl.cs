using System.IO.Compression;
using System.Xml.Linq;

namespace MergeDocuments.Services
{
    public class UpdateContentControl
    {
        // Finds content controls by their tag value and replaces their content with the given text.
        public void UpdateContentControls(string mergedFilePath, Dictionary<string, string> replacements)
        {
            if (string.IsNullOrEmpty(mergedFilePath))
                throw new ArgumentException("Merged file path cannot be null or empty", nameof(mergedFilePath));
            if (!File.Exists(mergedFilePath))
                throw new FileNotFoundException("Merged file not found", mergedFilePath);
            if (replacements == null)
                throw new ArgumentNullException(nameof(replacements));
            if (replacements.Count == 0)
                throw new ArgumentException("Replacements dictionary cannot be empty", nameof(replacements));
            if (replacements.Keys.Any(string.IsNullOrEmpty))
                throw new ArgumentException("Replacement keys cannot be null or empty", nameof(replacements));
            if (replacements.Values.Any(v => v == null))
                throw new ArgumentException("Replacement values cannot be null", nameof(replacements));

            // Open the DOCX file as a ZIP archive for modification
            using var updateZip = ZipFile.Open(mergedFilePath, ZipArchiveMode.Update);

            // Load the main document XML
            var docEntry = updateZip.GetEntry("word/document.xml")!;
            var updateDoc = LoadXml(docEntry);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Iterate over all structured document tags (content controls)
            foreach (var sdt in updateDoc.Descendants(w + "sdt"))
            {
                // Get the tag identifier(w: tag w: val)
                var tag = sdt.Descendants(w + "tag").FirstOrDefault()?.Attribute(w + "val")?.Value;

                // If tag matches a replacement key
                if (tag != null && replacements.ContainsKey(tag))
                {
                    var sdtContent = sdt.Element(w + "sdtContent");
                    if (sdtContent != null)
                    {
                        // Remove existing nodes and insert new paragraph with replacement text
                        sdtContent.RemoveNodes();
                        sdtContent.Add(new XElement(w + "p",
                            new XElement(w + "r",
                                new XElement(w + "t",
                                    // Preserve spaces in text
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    replacements[tag]
                                )
                            )
                        ));
                    }
                }
            }

            // Save updated XML back into the DOCX package
            SaveXml(updateZip, "word/document.xml", updateDoc);
        }

        // Loads XML content from a ZIP archive entry.
        private static XDocument LoadXml(ZipArchiveEntry entry)
        {
            using var reader = new StreamReader(entry.Open());
            return XDocument.Load(reader);
        }

        // Saves XML content into a ZIP archive entry, replacing the old one if it exists.
        private static void SaveXml(ZipArchive zip, string path, XDocument doc)
        {
            zip.GetEntry(path)?.Delete();
            var newEntry = zip.CreateEntry(path);
            using var writer = new StreamWriter(newEntry.Open());
            doc.Save(writer);
        }
    }
}