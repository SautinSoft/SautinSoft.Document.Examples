using System;
using System.IO;
using System.Xml;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            SaveToXmlFile();
        }

        /// <summary>
        /// Creates a new document and saves it as Xml file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-xml-net-csharp-vb.php
        /// </remarks>
        static void SaveToXmlFile()
        {
            string inpFile = @"..\..\..\example.docx"; 
            string outFile = Path.ChangeExtension(inpFile, ".xml");

            var dc = DocumentCore.Load(inpFile);
            var xml = GetXml(dc);
            xml.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        private static XmlDocument GetXml(DocumentCore dc)
        {
            var xml = new XmlDocument();
            var body = NewXmlNode(xml, xml, XmlNodeType.Element, "Document");

            foreach (Section section in dc.Sections)
            {
                var sec = NewXmlNode(xml, body, XmlNodeType.Element, "Section");
                WriteBlock(xml, sec, section.Blocks);
            }
            return xml;
        }

        private static void WriteBlock(XmlDocument xml, XmlNode parent, BlockCollection blocks)
        {
            foreach (var block in blocks)
            {
                switch (block)
                {
                    case Paragraph paragraph:
                        if (paragraph.Inlines.Count > 0)
                        {
                            var par = NewXmlNode(xml, parent, XmlNodeType.Element, "Paragraph");
                            foreach (var line in paragraph.Inlines)
                            {
                                switch (line)
                                {
                                    case Run run:
                                        var runNode = NewXmlNode(xml, par, XmlNodeType.Element, "Run");
                                        runNode.InnerText = run.Text;
                                        break;

                                    case ShapeBase shape:
                                        WriteDrawing(xml, par, shape);
                                        break;
                                }
                            }
                        }
                        break;
                    case Table table:
                        WriteTable(xml, parent, table);
                        break;
                }
            }
        }

        private static void WriteDrawing(XmlDocument xml, XmlNode parent, Element item)
        {
            if (item is ShapeGroup)
            {
                var shape = NewXmlNode(xml, parent, XmlNodeType.Element, "ShapeGroup");
                foreach (var sh in (item as ShapeGroup).ChildShapes)
                {
                    WriteDrawing(xml, shape, sh);
                }
            }
            else if (item is Picture)
            {
                var picture = NewXmlNode(xml, parent, XmlNodeType.Element, "Picture");
                var attr = xml.CreateAttribute("name");
                attr.Value = (item as Picture).Description;
                picture.Attributes.Append(attr);
            }
            else
            {
                var shape = NewXmlNode(xml, parent, XmlNodeType.Element, "Shape");
                var attr = xml.CreateAttribute("Figure");
                attr.Value = (item as Shape).Geometry.IsPreset ? ((item as Shape).Geometry as PresetGeometry).Figure.ToString() : "Custom";
                shape.Attributes.Append(attr);
                WriteBlock(xml, shape, (item as Shape).Text.Blocks);
            }
        }

        private static void WriteTable(XmlDocument xml, XmlNode parent, Table table)
        {
            var tab = NewXmlNode(xml, parent, XmlNodeType.Element, "Table");
            XmlAttribute attr = xml.CreateAttribute("rows");
            attr.Value = table.Rows.Count.ToString();
            tab.Attributes.Append(attr);
            attr = xml.CreateAttribute("cols");
            attr.Value = table.Columns.Count.ToString();
            tab.Attributes.Append(attr);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = NewXmlNode(xml, tab, XmlNodeType.Element, "Row");
                for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                {
                    var cell = NewXmlNode(xml, row, XmlNodeType.Element, "Cell");
                    if (table.Rows[i].Cells[j].RowSpan > 1)
                    {
                        attr = xml.CreateAttribute("rowspan");
                        attr.Value = table.Rows[i].Cells[j].RowSpan.ToString();
                        cell.Attributes.Append(attr);
                    }
                    if (table.Rows[i].Cells[j].ColumnSpan > 1)
                    {
                        attr = xml.CreateAttribute("colspan");
                        attr.Value = table.Rows[i].Cells[j].ColumnSpan.ToString();
                        cell.Attributes.Append(attr);
                    }
                    WriteBlock(xml, cell, table.Rows[i].Cells[j].Blocks);
                }
            }
        }

        private static XmlNode NewXmlNode(XmlDocument xml, XmlNode parent, XmlNodeType type, string name)
        {
            XmlNode node = xml.CreateNode(type, name, null);
            parent.AppendChild(node);
            return node;
        }
    }
}