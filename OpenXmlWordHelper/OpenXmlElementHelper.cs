using DocumentFormat.OpenXml;

namespace OpenXmlWordHelper
{
    public static class OpenXmlElementHelper
    {
        public static bool HasChild<T>(OpenXmlElement parent) where T : OpenXmlElement => parent.GetFirstChild<T>() != null;

        public static void SetChild<T>(OpenXmlElement parent, T child) where T : OpenXmlElement
        {
            parent.RemoveAllChildren<T>();
            parent.AppendChild(child);
        }

        public static void SetChildIfNotExists<T>(OpenXmlElement parent) where T : OpenXmlElement, new()
        {
            if (!HasChild<T>(parent))
            {
                parent.AppendChild(new T());
            }
        }
    }
}
