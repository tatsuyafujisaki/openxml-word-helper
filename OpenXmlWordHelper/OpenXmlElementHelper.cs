using DocumentFormat.OpenXml;

namespace OpenXmlWordHelper
{
    static class OpenXmlElementHelper
    {
        internal static bool HasChild<T>(OpenXmlElement parent) where T : OpenXmlElement => parent.GetFirstChild<T>() != null;

        internal static void SetChild<T>(OpenXmlElement parent, T child) where T : OpenXmlElement
        {
            parent.RemoveAllChildren<T>();
            parent.AppendChild(child);
        }

        internal static void AppendChildIfNotExists<T>(OpenXmlElement parent) where T : OpenXmlElement, new()
        {
            if (!HasChild<T>(parent))
            {
                parent.AppendChild(new T());
            }
        }
    }
}
