using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;

namespace OpenXmlWordHelper
{
    static class ParagraphPropertiesHelper
    {
        internal static void SetParagraphProperties(ParagraphProperties p, int? leftChars, int? firstLineChars, int? hangingChars, int? spaceBetweenLines, JustificationValues? jv)
        {
            var oxes = new List<OpenXmlElement>();

            if (leftChars != null || firstLineChars != null || hangingChars != null)
            {
                var indent = new Indentation();

                if (leftChars != null)
                {
                    indent.LeftChars = leftChars.Value;
                }

                if (firstLineChars != null)
                {
                    indent.FirstLineChars = firstLineChars.Value;
                }

                if (hangingChars != null)
                {
                    indent.HangingChars = hangingChars.Value;
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

        internal static void SetIndent(ParagraphProperties pp, int? leftChars, int? firstLineChars, int? hangingChars)
        {
            var indent = new Indentation();

            if (leftChars != null)
            {
                indent.LeftChars = leftChars.Value;
            }

            if (firstLineChars != null)
            {
                indent.FirstLineChars = firstLineChars.Value;
            }

            if (hangingChars != null)
            {
                indent.HangingChars = hangingChars.Value;
            }

            OpenXmlElementHelper.SetChild(pp, indent);
        }

        internal static void SetSpacinggBetweenLines(ParagraphProperties pp, int space)
        {
            OpenXmlElementHelper.SetChild(pp, new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto, Line = space.ToString() });
        }

        internal static void SetJustification(ParagraphProperties pp, JustificationValues jv)
        {
            OpenXmlElementHelper.SetChild(pp, new Justification { Val = jv });
        }
    }
}
