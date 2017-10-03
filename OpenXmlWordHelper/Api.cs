using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Math = DocumentFormat.OpenXml.Math;

namespace OpenXmlWordHelper
{
    [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
    [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
    public static class Api
    {
        public static class Equation
        {
            public static Paragraph CreateEquation(string equation) =>
                            new Paragraph(
                                new Math.Paragraph(
                                    new Math.OfficeMath(
                                        new Math.Run(
                                            new Math.RunProperties(
                                                new Math.Text { Text = equation })))));

            public static Paragraph CreateEquationInSpecifiedFont(string equation, string fontName) =>
                new Paragraph(
                    new Math.Paragraph(
                        new Math.OfficeMath(
                            new Math.Run(
                                new Math.RunProperties(new Math.NormalText()),
                                new RunFonts { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName, ComplexScript = fontName },
                                new Math.Text { Text = equation }))));

        }

        public static void ProtectWord(string path, string password)
        {
            string CreateSalt()
            {
                const int saltSize = 8;
                var salt = new byte[saltSize];
                RandomNumberGenerator.Create().GetNonZeroBytes(salt);
                return Convert.ToBase64String(salt);
            }

            const int spinCount = 100;

            string CreateHash(string salt)
            {
                byte[] ComputeHash(HashAlgorithm ha, byte[] a, byte[] b) => ha.ComputeHash(a.Concat(b).ToArray());

                var algorithm = new SHA512Managed();
                var hash = ComputeHash(algorithm, Convert.FromBase64String(salt), Encoding.Unicode.GetBytes(password));
                var bytes = new byte[4];

                for (var i = 0; i < spinCount; i++)
                {
                    Array.Copy(BitConverter.GetBytes(i), bytes, bytes.Length);
                    hash = ComputeHash(algorithm, hash, bytes);
                }

                return Convert.ToBase64String(hash);
            }

            // https://msdn.microsoft.com/library/documentformat.openxml.wordprocessing.writeprotection.cryptographicalgorithmsid.aspx
            const int sha512 = 14;

            using (var wd = WordprocessingDocument.Open(path, true))
            {
                var salt = CreateSalt();

                var dp = new DocumentProtection
                {
                    Edit = DocumentProtectionValues.ReadOnly,
                    CryptographicAlgorithmSid = sha512,
                    CryptographicSpinCount = spinCount,
                    Salt = salt,
                    Hash = CreateHash(salt)
                };

                var dsp = wd.MainDocumentPart.DocumentSettingsPart;

                var oldDp = dsp.Settings.FirstOrDefault(s => s.GetType() == typeof(DocumentProtection));

                if (oldDp == null)
                {
                    dsp.Settings.AppendChild(dp);
                }
                else
                {
                    dsp.Settings.ReplaceChild(dp, oldDp);
                }

                dsp.Settings.Save();
            }
        }

        static void AddStyleDefinitionPart(OpenXmlPartContainer oxpc) => new Styles().Save(oxpc.AddNewPart<StyleDefinitionsPart>());

        public static void AddParagraphPropertiesDefault(MainDocumentPart mdp)
        {
            if (mdp.StyleDefinitionsPart == null)
            {
                AddStyleDefinitionPart(mdp);
            }

            mdp.StyleDefinitionsPart.Styles = new Styles(new DocDefaults(new ParagraphPropertiesDefault()));
        }

        public static void ClearHeaderFooter(MainDocumentPart mdp)
        {
            mdp.DeleteParts(mdp.HeaderParts);
            mdp.DeleteParts(mdp.FooterParts);

            var hp = mdp.AddNewPart<HeaderPart>();
            var fp = mdp.AddNewPart<FooterPart>();

            hp.Header = new Header();
            fp.Footer = new Footer();

            foreach (var sps in mdp.Document.Body.Elements<SectionProperties>())
            {
                sps.RemoveAllChildren<HeaderReference>();
                sps.RemoveAllChildren<FooterReference>();
                sps.PrependChild(new HeaderReference { Id = mdp.GetIdOfPart(hp) });
                sps.PrependChild(new FooterReference { Id = mdp.GetIdOfPart(fp) });
            }
        }

        public static void MergeDocuments(string sourcePath1, string sourcePath2)
        {
            using (var wd = WordprocessingDocument.Open(sourcePath1, true))
            {
                const string id = "altCunkt1";

                var mdp = wd.MainDocumentPart;

                using (var fs = File.OpenRead(sourcePath2))
                {
                    mdp.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, id).FeedData(fs);
                }

                mdp.Document.Body.InsertAfter(new AltChunk { Id = id }, mdp.Document.Body.Elements<Paragraph>().Last());
            }
        }

        public static void MergeDocuments(string sourcePath1, string sourcePath2, string destinationPath)
        {
            File.Delete(destinationPath);
            File.Copy(sourcePath1, destinationPath);

            MergeDocuments(destinationPath, sourcePath2);
        }

        public static AbstractNum CreateAbstractNum(int abstractNumId)
        {
            if (abstractNumId < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            var an = new AbstractNum { AbstractNumberId = abstractNumId };

            an.AppendChild(LevelFactory.CreateLevel(NumberFormatValues.Decimal, 720, 360));
            an.AppendChild(LevelFactory.CreateLevel(NumberFormatValues.Decimal, 720 * 2, 360));
            an.AppendChild(LevelFactory.CreateLevel(NumberFormatValues.Decimal, 720 * 3, 360));

            return an;
        }

        public static NumberingInstance CreateNumberingInstance(int numberId, int abstractNumId)
        {
            if (numberId < 1)
            {
                throw new ArgumentOutOfRangeException();
            }

            if (abstractNumId < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            return new NumberingInstance
            {
                NumberID = numberId,
                AbstractNumId = new AbstractNumId { Val = abstractNumId }
            };
        }

        public static Paragraph CreateNumberingParagraph(int numberId, int levelIndex, string s)
        {
            // Note
            // NumberingLevelRefrence.Val == Level.LevelIndex

            if (numberId < 1)
            {
                throw new ArgumentOutOfRangeException();
            }

            if (levelIndex < 0 || 2 < levelIndex)
            {
                throw new ArgumentOutOfRangeException();
            }

            var pps = new ParagraphProperties(new NumberingProperties(new NumberingId { Val = numberId }, new NumberingLevelReference { Val = levelIndex }));

            var run = new Run(new Text(s));

            return new Paragraph(pps, run);
        }

        static void DeletePart<T>(WorksheetPart wp) where T : OpenXmlPart
        {
            var parts = wp.GetPartsOfType<T>().ToList();

            if (parts.Any())
            {
                foreach (var part in parts)
                {
                    wp.DeletePart(part);
                }
            }
        }

        // Colunn index is zero-based.
        public static void SetColumnJustification(Table table, int columnIndex, JustificationValues jv)
        {
            foreach (var tr in table.Elements<TableRow>())
            {
                var tc = columnIndex == 0 ? tr.GetFirstChild<TableCell>() : tr.Elements<TableCell>().ToList()[columnIndex];

                var p = tc.GetFirstChild<Paragraph>();

                OpenXmlElementHelper.SetChildIfNotExists<ParagraphProperties>(p);
                OpenXmlElementHelper.SetChild(p.GetFirstChild<ParagraphProperties>(), new Justification { Val = jv });
            }
        }
    }
}