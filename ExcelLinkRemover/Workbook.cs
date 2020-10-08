using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace ExcelLinkRemover
{
    /// <summary>
    /// Class for working with Workbook.
    /// </summary>
    internal class Workbook
    {
        private const string externalLinks = "\\xl\\externalLinks";
        private const string workbook = "\\xl\\workbook.xml";
        private const string workbook_rels = "\\xl\\_rels\\workbook.xml.rels";
        private const string worksheets = "\\xl\\worksheets";

        /// <summary>
        /// Removes all external links in unpacked files.
        /// </summary>
        /// <param name="tempPath">Full path to the temporary folder.</param>
        /// <returns>Returns the number of operations performed.</returns>
        internal static int Process(string tempPath)
        {
            var count = 0;
            var regex = new Regex(@"\[[0-9]+\]");
            var path = tempPath + workbook;
            var links = tempPath + externalLinks;
            var doc = new XmlDocument();
            doc.Load(path);
            var root = doc.DocumentElement;
            var names = root["definedNames"];
            var references = root["externalReferences"];

            // Processing all defined names.
            foreach (XmlNode name in names)
            {
                // If it contains an external link, we remove it.
                // [TODO]: The find and replace solution "[0-9]" with "" is not ideal and needs something better.
                if (regex.Matches(name.FirstChild.Value).Count > 0)
                {
                    name.FirstChild.Value = regex.Replace(name.FirstChild.Value, "");
                }
            }

            // Processing all external connections.
            if (references != null)
            {
                var ids = new string[references.ChildNodes.Count];

                for (int i = 0; i < ids.Length; i++)
                {
                    ids[i] = references.ChildNodes[i].Attributes.GetNamedItem("r:id")?.Value;
                }

                // Removing links to external relations.
                count += ProcessRels(tempPath, ids);

                // Removing the block with external references, if any.
                root.RemoveChild(references);
            }

            // Deleting the folder with external relations
            if (Directory.Exists(links))
            {
                try
                {
                    Directory.Delete(links, true);
                }
                catch
                { }
            }

            // Processing all Sheets.
            var directory = new DirectoryInfo(tempPath + worksheets);

            foreach (FileInfo file in directory.GetFiles())
            {
                count += Sheet.Process(file.FullName);
            }

            // Save if changes have been made.
            doc.Save(path);

            return count;
        }

        /// <summary>
        /// Removes links to external relation.
        /// </summary>
        /// <param name="tempPath">Full path to the temporary folder.</param>
        /// <param name="ids">List of identifiers.</param>
        /// <returns>Returns the number of operations performed.</returns>
        private static int ProcessRels(string tempPath, string[] ids)
        {
            var count = 0;
            var path = tempPath + workbook_rels;
            var doc = new XmlDocument();
            doc.Load(path);
            var root = doc.DocumentElement;

            // Processing all relations in the document.
            foreach (XmlNode relationship in root)
            {
                foreach (string id in ids)
                {
                    // Remove link to external relation.
                    if (relationship.Attributes.GetNamedItem("Id")?.Value == id)
                    {
                        root.RemoveChild(relationship);
                        count++;
                    }
                }
            }

            // Save if changes have been made.
            if (count > 0)
            {
                doc.Save(path);
            }

            return count;
        }
    }
}
