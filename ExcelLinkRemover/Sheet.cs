using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelLinkRemover
{
    /// <summary>
    /// Class for working with Sheet.
    /// </summary>
    public class Sheet
    {
        /// <summary>
        /// Removes all external links on the current Sheet.
        /// </summary>
        /// <param name="fileName">Full path to Sheet.xml file.</param>
        /// <returns>Returns the number of operations performed.</returns>
        internal static int Process(string fileName)
        {
            var count = 0;
            var regex = new Regex(@"\[[0-9]+\]");
            var doc = new XmlDocument();
            doc.Load(fileName);
            var root = doc.DocumentElement;
            var sheet = root["sheetData"];
            var extLst = root["extLst"];

            // Removing the block with external links, if any.
            if (extLst != null)
            {
                root.RemoveChild(extLst);
            }

            // [OLD] Old solution
            // Remove xrefs in each cell, if any.
            //foreach (XmlNode row in sheet)
            //{
            //    foreach (XmlNode cell in row)
            //    {
            //        var attr = cell.Attributes.GetNamedItem("t");

            //        if (attr != null && attr.Value == "str")
            //        {
            //            foreach (XmlNode item in cell)
            //            {
            //                [TODO]:
            //                // 1.The condition item.Attributes.Count == 0 requires checking for more cases.
            //                // 2.The solution to find and replace "[0-9]" with "" is not ideal and requires something better.
            //                if (item.Name == "f" && item.Attributes.Count == 0 && regex.Matches(item.FirstChild.Value).Count > 0)
            //                {
            //                    item.FirstChild.Value = regex.Replace(item.FirstChild.Value, "");
            //                    count++;
            //                }
            //            }
            //        }
            //    }
            //}

            // Removing xrefs in each cell, if any.
            // We use parallel threads to speed up processing.
            var result = Parallel.ForEach(sheet.Cast<XmlNode>(),
            item3 =>
            {
                var row = item3.ChildNodes;
                foreach (XmlNode cell in row)
                {
                    var attr2 = cell.Attributes.GetNamedItem("t");

                    if (attr2 != null && attr2.Value == "str")
                    {
                        foreach (XmlNode item2 in cell)
                        {
                            // [TODO]:
                            // 1.The condition item.Attributes.Count == 0 requires checking for more cases.
                            // 2.The solution to find and replace "[0-9]" with "" is not ideal and requires something better.
                            if (item2.Name == "f" && item2.Attributes.Count == 0 && regex.Matches(item2.FirstChild.Value).Count > 0)
                            {
                                item2.FirstChild.Value = regex.Replace(item2.FirstChild.Value, "");
                                count++;
                            }
                        }
                    }
                }
            });

            // Save if changes have been made.
            if (count > 0)
            {
                doc.Save(fileName);
            }

            return count;
        }
    }
}