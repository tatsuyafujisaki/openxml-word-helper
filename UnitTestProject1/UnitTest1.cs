using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlWordHelper;
using System;
using System.Diagnostics;
using System.IO;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [ClassInitialize]
        public static void Initialize(TestContext context) => Environment.CurrentDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

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

        static void TestRunner(string path, Action<MainDocumentPart> f)
        {
            using (var wd = WordprocessingDocument.Open(path, true))
            {
                f(wd.MainDocumentPart);
            }

            Process.Start(path);
        }

        [TestMethod]
        public void TestCreateParagraphProperties() =>
            TestRunner(mdp =>
            {
                var pp = new ParagraphProperties();
                ParagraphPropertiesHelper.SetParagraphProperties(pp, 10, new FirstLineOrHangingChars(FirstLineOrHanging.FirstLine, 10), 1000, JustificationValues.Both);
                var p = new Paragraph();
                OpenXmlElementHelper.SetChild(p, pp);
                OpenXmlElementHelper.SetChild(p, new Run(new Text("Hello world!")));
                OpenXmlElementHelper.SetChild(mdp.Document.Body, p);
            });


        [TestMethod]
        public void TestAddParagraphPropertiesDefault() => TestRunner(Api.AddParagraphPropertiesDefault);

        [TestMethod]
        public void TestCreateWordWithParagraphPropertiesDefault() => TestRunner(Api.ClearHeaderFooter);

        [TestMethod]
        public void TestCreateNumberingParagraphs()
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
        public void TestMergeDocuments()
        {
            const string path1 = "Source1.docx";
            const string path2 = "Source2.docx";

            Api.MergeDocuments(path1, path2);

            Process.Start(path1);
        }

        [TestMethod]
        public void TestMergeDocumentsToNewFile()
        {
            const string path1 = "Source1.docx";
            const string path2 = "Source2.docx";
            const string path3 = "Destination.docx";

            Api.MergeDocuments(path1, path2, path3);

            Process.Start(path3);
        }

        [TestMethod]
        public void TestProtectWord()
        {
            const string path = "Sample.docx";

            Api.ProtectWord(path, "dummy");

            Process.Start(path);
        }

        [TestMethod]
        public void TestSetColumnJustification()
        {
            const string path = "Sample.docx";

            TestRunner(path, mdp => Api.SetColumnJustification(mdp.Document.Body.GetFirstChild<Table>(), 0, JustificationValues.Center));
        }

        [TestMethod]
        public void TestCreateEquation() =>
            TestRunner(mdp =>
            {
                mdp.Document.Body.AppendChild(Api.Equation.CreateEquation("1 + 1 = 2"));
            });

        [TestMethod]
        public void TestCreateEquationInSpecifiedFont() =>
            TestRunner(mdp =>
            {
                mdp.Document.Body.AppendChild(Api.Equation.CreateEquationInSpecifiedFont("1 + 1 = 2", "Yu Gothic UI"));
            });
    }
}