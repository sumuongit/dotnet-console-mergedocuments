using System.IO.Compression;
using System.Xml.Linq;

namespace MergeDocuments.Services
{
    public class DynamicFooter
    {
        // Creates and adds a footer XML part with dynamic page numbering for a specific document.
        public void CreateFooter(ZipArchive zip, XDocument relsDoc, XDocument contentTypesDoc,
                                 string filePath, string relId, string footerFileName)
        {
            // Build the footer XML containing the filename and page number field
            var footerXml = BuildFooterXDocument(Path.GetFileName(filePath));

            // Remove existing footer part if present
            zip.GetEntry("word/" + footerFileName)?.Delete();

            // Add new footer part to the archive
            var newFooterEntry = zip.CreateEntry("word/" + footerFileName);
            using (var wtr = new StreamWriter(newFooterEntry.Open()))
                footerXml.Save(wtr);

            // Add a relationship entry for the new footer in document.xml.rels
            relsDoc.Root.Add(new XElement(relsDoc.Root.Name.Namespace + "Relationship",
                new XAttribute("Id", relId),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                new XAttribute("Target", footerFileName)));

            // Ensure content types file has an override entry for this footer part
            bool ctExists = contentTypesDoc.Root.Elements()
                .Any(e => e.Name.LocalName == "Override" && (string)e.Attribute("PartName") == "/word/" + footerFileName);

            if (!ctExists)
            {
                contentTypesDoc.Root.Add(new XElement(contentTypesDoc.Root.Name.Namespace + "Override",
                    new XAttribute("PartName", "/word/" + footerFileName),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));
            }
        }

        // Builds the footer XML document containing the filename and a Word PAGE field for dynamic page numbers
        private static XDocument BuildFooterXDocument(string fileName)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            var ftr = new XElement(w + "ftr",
                new XAttribute(XNamespace.Xmlns + "w", w.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", r.NamespaceName),

                // Paragraph containing "<filename> - Page <PAGE>"
                new XElement(w + "p",
                    new XElement(w + "r", new XElement(w + "t", fileName + " - Page ")),
                    new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin"))),
                    new XElement(w + "r", new XElement(w + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), "PAGE")),
                    new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end")))
                )
            );

            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), ftr);
        }
    }
}