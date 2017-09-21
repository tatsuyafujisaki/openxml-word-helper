using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using Italic = DocumentFormat.OpenXml.Wordprocessing.Italic;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using UnderlineValues = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues;

namespace OpenXmlWordHelper
{
    static class RuPropertiesHelper
    {
        internal static void SetRunProperties(Run run, bool isBold, bool isItalic, bool isUnderline, string fontName, int? fontSize)
        {
            var oxes = new List<OpenXmlElement>();

            if (isBold)
            {
                oxes.Add(new Bold());
            }

            if (isItalic)
            {
                oxes.Add(new Italic());
            }

            if (isUnderline)
            {
                oxes.Add(new Underline { Val = UnderlineValues.Single });
            }

            if (fontName != null)
            {
                oxes.Add(new RunFonts
                {
                    Ascii = fontName,
                    HighAnsi = fontName,
                    EastAsia = fontName,
                    ComplexScript = fontName
                });
            }

            if (fontSize.HasValue)
            {
                oxes.Add(new FontSize { Val = fontSize.Value.ToString() });
                oxes.Add(new FontSizeComplexScript { Val = fontSize.Value.ToString() });
            }

            OpenXmlElementHelper.SetChild(run, new RunProperties(oxes));
        }

        internal static void SetBold(RunProperties rp)
        {
            OpenXmlElementHelper.SetChild(rp, new Bold());
        }

        internal static void SetItalic(RunProperties rp)
        {
            OpenXmlElementHelper.SetChild(rp, new Italic());
        }

        internal static void SetUnderline(RunProperties rp)
        {
            OpenXmlElementHelper.SetChild(rp, new Underline { Val = UnderlineValues.Single });
        }

        internal static void SetFontName(RunProperties rp, string fontName)
        {
            OpenXmlElementHelper.SetChild(rp, new RunFonts
            {
                Ascii = fontName,
                HighAnsi = fontName,
                EastAsia = fontName,
                ComplexScript = fontName
            });
        }

        internal static void SetFontSize(RunProperties rp, int fontSize)
        {
            OpenXmlElementHelper.SetChild(rp, new FontSize { Val = fontSize.ToString() });
            OpenXmlElementHelper.SetChild(rp, new FontSizeComplexScript { Val = fontSize.ToString() });
        }
    }
}
