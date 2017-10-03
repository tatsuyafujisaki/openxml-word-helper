using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.ComponentModel;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;

namespace OpenXmlWordHelper
{
    public enum FirstLineOrHanging
    {
        FirstLine,
        Hanging
    }

    public class FirstLineOrHangingChars
    {
        internal readonly FirstLineOrHanging Floh;
        internal readonly int Chars;

        public FirstLineOrHangingChars(FirstLineOrHanging floh, int chars)
        {
            this.Floh = floh;
            this.Chars = chars;
        }
    }

    public static class ParagraphPropertiesHelper
    {
        public static void SetParagraphProperties(ParagraphProperties p, int? leftChars, FirstLineOrHangingChars firstLineOrhangingChars, int? spaceBetweenLines, JustificationValues? jv)
        {
            var oxes = new List<OpenXmlElement>();

            if (leftChars != null || firstLineOrhangingChars != null)
            {
                var indent = new Indentation();

                if (leftChars != null)
                {
                    indent.LeftChars = leftChars.Value;
                }

                if (firstLineOrhangingChars != null)
                {
                    switch (firstLineOrhangingChars.Floh)
                    {
                        case FirstLineOrHanging.FirstLine:
                            indent.FirstLineChars = firstLineOrhangingChars.Chars;
                            break;
                        case FirstLineOrHanging.Hanging:
                            indent.HangingChars = firstLineOrhangingChars.Chars;
                            break;
                        default:
                            throw new InvalidEnumArgumentException(firstLineOrhangingChars.Floh.ToString());
                    }
                }

                oxes.Add(indent);
            }

            if (spaceBetweenLines.HasValue)
            {
                oxes.Add(new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto, Line = spaceBetweenLines.Value.ToString() });
            }

            if (jv.HasValue)
            {
                oxes.Add(new Justification { Val = jv });
            }

            OpenXmlElementHelper.SetChild(p, new ParagraphProperties(oxes));
        }

        public static void SetIndent(ParagraphProperties pp, int? leftChars, FirstLineOrHangingChars firstLineOrhangingChars)
        {
            var indent = new Indentation();

            if (leftChars != null)
            {
                indent.LeftChars = leftChars.Value;
            }

            if (firstLineOrhangingChars != null)
            {
                switch (firstLineOrhangingChars.Floh)
                {
                    case FirstLineOrHanging.FirstLine:
                        indent.FirstLineChars = firstLineOrhangingChars.Chars;
                        break;
                    case FirstLineOrHanging.Hanging:
                        indent.HangingChars = firstLineOrhangingChars.Chars;
                        break;
                    default:
                        throw new InvalidEnumArgumentException(firstLineOrhangingChars.Floh.ToString());
                }
            }

            OpenXmlElementHelper.SetChild(pp, indent);
        }

        public static void SetSpacinggBetweenLines(ParagraphProperties pp, int space)
        {
            OpenXmlElementHelper.SetChild(pp, new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto, Line = space.ToString() });
        }

        public static void SetJustification(ParagraphProperties pp, JustificationValues jv)
        {
            OpenXmlElementHelper.SetChild(pp, new Justification { Val = jv });
        }
    }
}