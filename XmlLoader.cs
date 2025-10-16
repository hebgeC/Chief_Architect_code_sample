namespace SpreadsheetEngine
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Xml;

    /// <summary>
    /// Class to load a spreadsheet from an XML file.
    /// </summary>
    internal static class XmlLoader
    {
        /// <summary>
        /// Load a spreadsheet from an XML file.
        /// </summary>
        /// <param name="grid">spreadsheet cells.</param>
        /// <param name="fs">filestream.</param>
        public static void Load(Cell[,] grid, FileStream fs)
        {
            // clear the spreadsheet
            foreach (Cell cell in grid)
            {
                cell.Text = string.Empty;
                cell.BGColor = 0xFFFFFFFF;
            }

            XmlReaderSettings settings = new XmlReaderSettings
            {
                IgnoreWhitespace = true,
                IgnoreComments = true,
            };



            using (XmlReader reader = XmlReader.Create(fs, settings))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement() && reader.Name == "cell")
                    {
                        string? rowValue = reader["row"];
                        string? colValue = reader["col"];

                        if (string.IsNullOrEmpty(rowValue) || string.IsNullOrEmpty(colValue))
                        {
                            throw new InvalidOperationException("cell has invalid indexing");
                        }

                        int rowIndex = int.Parse(rowValue);
                        int columnIndex = int.Parse(colValue);
                        string cellText = string.Empty;
                        string bgColor = string.Empty;

                        // now reading elements within the <cell> element
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "cell")
                            {
                                break;
                            }

                            if (reader.IsStartElement() && reader.Name == "text")
                            {
                                reader.Read(); // reading whats inside the <text> element
                                cellText = reader.Value;
                            }

                            if (reader.IsStartElement() && reader.Name == "bgColor")
                            {
                                reader.Read();
                                bgColor = reader.Value;
                            }
                        }

                        if (!string.IsNullOrEmpty(cellText))
                        {
                            grid[rowIndex, columnIndex].Text = cellText;
                        }

                        if (!string.IsNullOrEmpty(bgColor))
                        {
                            grid[rowIndex, columnIndex].BGColor = uint.Parse(bgColor, System.Globalization.NumberStyles.HexNumber);
                        }
                    }
                }
            }
        }
    }
}
