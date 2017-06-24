using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlWordHelper;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace Test1
{
    [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
    [TestClass]
    public class UnitTest1
    {
        static void TestRunner(Action<MainDocumentPart> f)
        {
            var path = Path.Combine(Path.GetTempPath(), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss-fffffff") + ".docx");

            using (var wd = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var mdp = wd.AddMainDocumentPart();
                mdp.Document = new Document(new Body());

                f(mdp);
            }

            Process.Start(path);
        }

        [TestMethod]
        public static void TestAddParagraphPropertiesDefault() => TestRunner(Api.AddParagraphPropertiesDefault);

        [TestMethod]
        public static void TestCreateWordWithParagraphPropertiesDefault() => TestRunner(Api.ClearHeaderFooter);

        [TestMethod]
        public static void TestCreateParagraphWithText()
        {
            TestRunner(mdp =>
            {
                mdp.Document.Body.AppendChild(Api.CreateParagraphWithText("Hello"));
                mdp.Document.Body.AppendChild(Api.CreateParagraphWithText("Word"));
            });
        }

        [TestMethod]
        public static void TestCreateNumberingParagraphs()
        {
            TestRunner(mdp =>
            {
                const int abstractNumId = 0;
                const int numberId = 2;

                mdp.Document = new Document(new Body());

                mdp.AddNewPart<StyleDefinitionsPart>();
                mdp.StyleDefinitionsPart.Styles = new Styles(new DocDefaults(new ParagraphPropertiesDefault()));

                mdp.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering();

                mdp.NumberingDefinitionsPart.Numbering.Append(Api.CreateAbstractNum(abstractNumId));
                mdp.NumberingDefinitionsPart.Numbering.Append(Api.CreateNumberingInstance(numberId, abstractNumId));

                mdp.Document.Body.Append(Api.CreateNumberingParagraph(numberId, 0, "Foo"));
                mdp.Document.Body.Append(Api.CreateNumberingParagraph(numberId, 1, "Bar"));
                mdp.Document.Body.Append(Api.CreateNumberingParagraph(numberId, 2, "Baz"));
            });
        }

        [TestMethod]
        public static void TestMergeDocuments()
        {
            const string path1 = "Source1.docx";
            const string path2 = "Source2.docx";

            Api.MergeDocuments(path1, path2);

            Process.Start(path1);
        }

        [TestMethod]
        public static void TestMergeDocumentsToNewFile()
        {
            const string path1 = "Source1.docx";
            const string path2 = "Source2.docx";
            const string path3 = "Destination.docx";

            Api.MergeDocuments(path1, path2, path3);

            Process.Start(path3);
        }

        [TestMethod]
        public static void TestProtectWord()
        {
            const string path = "Sample.docx";

            Api.ProtectWord(path, "dummy");

            Process.Start(path);
        }
    }
}