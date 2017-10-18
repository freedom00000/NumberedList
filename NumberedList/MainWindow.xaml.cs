using System;
using System.Collections.Generic;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace NumberedList
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        string filepath = @"test.docx";

        public static XmlDocument GetXmlDocument(OpenXmlPart part)
        {
            XmlDocument xmlDoc = new XmlDocument();
            using (Stream partStream = part.GetStream())
            using (XmlReader partXmlReader = XmlReader.Create(partStream))
                xmlDoc.Load(partXmlReader);
            return xmlDoc;
        }

        public static void PutXmlDocument(OpenXmlPart part, XmlDocument xmlDoc)
        {
            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                xmlDoc.Save(partXmlWriter);
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            using(WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, true))
            {

                XmlDocument xmlDoc;
                xmlDoc = GetXmlDocument(wordDoc.MainDocumentPart);
                SearchAndReplaceInXmlDocument(xmlDoc);
                PutXmlDocument(wordDoc.MainDocumentPart, xmlDoc);
                wordDoc.Close();
            }
        }


        private static void SearchAndReplaceInXmlDocument(XmlDocument xmlDocument)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDocument.NameTable);
            nsmgr.AddNamespace("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var paragraphs = xmlDocument.SelectNodes("descendant::w:p", nsmgr);
            foreach (var paragraph in paragraphs)
                SearchAndReplaceInParagraph((XmlElement)paragraph);
        }

        private static void SearchAndReplaceInParagraph(XmlElement paragraph)
        {
            XmlDocument xmlDoc = paragraph.OwnerDocument;

            string wordNamespace =
               "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", wordNamespace);
            XmlNodeList paragraphNum = paragraph.SelectNodes("descendant::w:numPr", nsmgr);
            XmlNodeList paragraphText = paragraph.SelectNodes("descendant::w:t", nsmgr);

            StringBuilder sb = new StringBuilder();
            if (paragraphNum.Count != 0)
            {
                if (Char.IsUpper(paragraphText[0].InnerText[0]) 
                    && paragraphText[paragraphText.Count-1].InnerText.Contains("."))
                    paragraphNum[0].LastChild.Attributes["w:val"].Value = "10";
                else
                {
                    if (!Char.IsUpper(paragraphText[0].InnerText[0])
                        && paragraphText[paragraphText.Count - 1].InnerText.Contains(";"))
                        paragraphNum[0].LastChild.Attributes["w:val"].Value = "9";
                }
            }

        }
    }
}
