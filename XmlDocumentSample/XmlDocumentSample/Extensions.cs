using System.IO;
using System.Xml;
using System;


namespace XmlDocumentSample
{
    public static class Extensions
    {
        private static string ConvertToXPath(string input, string ns)
        {
            if (String.IsNullOrEmpty(ns) || String.IsNullOrEmpty(input))
                return input;

            var str = input.Split('/');
            string xPath = String.IsNullOrEmpty(str[0]) ? "" : ns + ":" + str[0];
            for (var i = 1; i < str.Length; i++)
            {
                xPath = String.Format("{0}/{1}:{2}", xPath, ns, str[i]);
            }

            return xPath.Trim();
        }

        public static string AsString(this XmlDocument xmlDoc)
        {
            using (StringWriter sw = new StringWriter())
            {
                using (XmlTextWriter tx = new XmlTextWriter(sw))
                {
                    xmlDoc.WriteTo(tx);
                    string strXmlText = sw.ToString();
                    return strXmlText;
                }
            }
        }

        public static XmlDocument RemoveChildrenAtPath(this XmlDocument doc, string path, string namespacePrefix, XmlNamespaceManager nsmgr)
        {
            string xPath = ConvertToXPath(path, namespacePrefix);
            XmlNodeList nodeList = doc.SelectNodes(xPath, nsmgr);

            foreach (XmlElement node in nodeList)
                node.ParentNode.RemoveChild(node);

            return doc;
        }

        public static XmlDocument RemoveAttributeAtPath(this XmlDocument doc, string path, string attrName, string namespacePrefix, XmlNamespaceManager nsmgr)
        {
            string xPath = ConvertToXPath(path, namespacePrefix);
            XmlNodeList nodeList = doc.SelectNodes(xPath, nsmgr);

            foreach (XmlElement node in nodeList)
                if (node.GetAttribute(attrName) != null)
                    node.RemoveAttribute(attrName);

            return doc;
        }
    }
}
