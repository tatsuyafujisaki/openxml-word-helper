using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OpenXmlWordHelper
{
    static class LevelFactory
    {
        static int _levelIndex = -1;

        internal static Level CreateLevel(NumberFormatValues nfv, int startIndent, int hangingIndent)
        {
            _levelIndex++;

            return new Level(new StartNumberingValue { Val = 1 },

                    new NumberingFormat { Val = nfv },
                    new LevelText { Val = $"%{_levelIndex + 1}." },
                    new PreviousParagraphProperties(new Indentation
                    {
                        Start = startIndent.ToString(CultureInfo.InvariantCulture),
                        Hanging = hangingIndent.ToString(CultureInfo.InvariantCulture)
                    }))
            { LevelIndex = _levelIndex };
        }
    }
}