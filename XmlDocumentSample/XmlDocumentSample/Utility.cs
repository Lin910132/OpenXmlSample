using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Globalization;
using System.Linq;

namespace XmlDocumentSample
{
    public class Utility
    {
        private const string _defaultPrefix = "default";

        #region Helpers
        private string ConvertToXPath(string input, string ns)
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

        private DateTime ParseDatetime(string value)
        {
            var valueDate = DateTime.MinValue;
            var valueDateTimeOffset = DateTimeOffset.MinValue;
            if (value != null)
            {
                if (!DateTime.TryParseExact(value,
                    new[]
                            {
                               "yyyyMMdd",
                                "yyyyMMddHHmmss"
                            },
                    CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out valueDate))
                {

                    if (!DateTime.TryParse(value, CultureInfo.CurrentCulture,
                         DateTimeStyles.AllowWhiteSpaces, out valueDate))
                    {
                        if (!DateTimeOffset.TryParseExact(value,
                            new[] { "yyyyMMddHHmmsszzz" },
                            CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out valueDateTimeOffset))
                        {
                            if (!DateTimeOffset.TryParse(value, CultureInfo.CurrentCulture,
                                DateTimeStyles.AllowWhiteSpaces, out valueDateTimeOffset))
                            {
                                throw new Exception(String.Format("Can't parse Encounter Date: {0}", value));
                            }
                        }
                    }
                }
            }

            if (valueDateTimeOffset != DateTimeOffset.MinValue)
            {
                valueDate = valueDateTimeOffset.LocalDateTime;
            }

            return valueDate;
        }

        private XmlNamespaceManager SetUpNamespaces(XmlDocument doc)
        {
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            nsmgr.AddNamespace(_defaultPrefix, "urn:hl7-org:v3");
            nsmgr.AddNamespace("voc", "urn:hl7-org:v3/voc");
            nsmgr.AddNamespace("sdtc", "urn:hl7-org:sdtc");
            return nsmgr;
        }
        #endregion

        #region Wokers

        private void AddTjcHeader(XmlDocument doc, XmlNamespaceManager nsmgr)
        {
            XmlElement root = doc.DocumentElement;
            XmlNodeList nodeList = root.SelectNodes(ConvertToXPath("templateId", _defaultPrefix), nsmgr);

            XmlNode position = (nodeList != null && nodeList.Count > 0) ? nodeList[nodeList.Count - 1] : null;
            XmlNode commentNode = doc.CreateNode(XmlNodeType.Comment, "TJC Specs", nsmgr.LookupNamespace(_defaultPrefix));
            commentNode.InnerText = "TJC-specific document OID";

            XmlElement elem1 = doc.CreateElement("templateId", nsmgr.LookupNamespace(_defaultPrefix));
            elem1.SetAttribute("root", "1.3.6.1.4.1.33895.1.5");
            elem1.SetAttribute("extension", "2015-06");

            if (position == null)
            {
                root.PrependChild(elem1);
                root.PrependChild(commentNode);
            }
            else
            {
                var oldElem = root.SelectSingleNode(ConvertToXPath("templateId[@root = '1.3.6.1.4.1.33895.1.5' and @extension = '2015-06']", _defaultPrefix), nsmgr);
                if (oldElem == null)
                {
                    root.InsertAfter(commentNode, position);
                    root.InsertAfter(elem1, commentNode);
                }

            }

        }

        private void AddVenderTrackingId(XmlDocument doc, XmlNamespaceManager nsmgr, Guid venderTrackingId)
        {
            XmlNode parent = doc.SelectSingleNode(ConvertToXPath("ClinicalDocument/recordTarget/patientRole", _defaultPrefix), nsmgr);
            XmlNodeList nodeList = parent.SelectNodes(ConvertToXPath("id", _defaultPrefix), nsmgr);

            foreach (XmlElement node in nodeList)
            {
                parent.RemoveChild(node);
            }

            XmlElement venderTracking = doc.CreateElement("id", nsmgr.LookupNamespace("default"));
            venderTracking.SetAttribute("root", "1.3.6.1.4.1.33895");
            venderTracking.SetAttribute("extension", venderTrackingId.ToString());

            parent.PrependChild(venderTracking);
        }

        private void RemovePhiData(XmlDocument doc, XmlNamespaceManager nsmgr)
        {
            //Patient Name
            doc.RemoveChildrenAtPath("ClinicalDocument/recordTarget/patientRole/patient/name", "default", nsmgr);
            doc.RemoveChildrenAtPath("ClinicalDocument/recordTarget/patientRole/telecom", "default", nsmgr);

            //Patient Address
            doc.RemoveChildrenAtPath("ClinicalDocument/recordTarget/patientRole/addr", "default", nsmgr);

            //Guardian
            doc.RemoveChildrenAtPath("ClinicalDocument/recordTarget/patientRole/patient/guardian", "default", nsmgr);

            //Birth Place
            doc.RemoveChildrenAtPath("ClinicalDocument/recordTarget/patientRole/patient/birthplace", "default", nsmgr);

            //Author information
            doc.RemoveChildrenAtPath("ClinicalDocument/author/assignedAuthor", "default", nsmgr);

            //LegalAuthenticator 
            doc.RemoveChildrenAtPath("ClinicalDocument/legalAuthenticator", "default", nsmgr);

        }

        private void AddHco(XmlDocument doc, string hco, XmlNamespaceManager nsmgr)
        {
            XmlElement parent = (XmlElement)doc.SelectSingleNode(ConvertToXPath("ClinicalDocument/custodian/assignedCustodian/representedCustodianOrganization", _defaultPrefix), nsmgr);
            var ccnElement = parent.SelectSingleNode(ConvertToXPath("id[@root = '2.16.840.1.113883.4.336']", _defaultPrefix), nsmgr);
            var hcoElement = parent.SelectSingleNode(ConvertToXPath("id[@root = '1.3.6.1.4.1.33895']", _defaultPrefix), nsmgr);

            XmlElement newHcoElem = doc.CreateElement("id", nsmgr.LookupNamespace(_defaultPrefix));
            newHcoElem.SetAttribute("root", "1.3.6.1.4.1.33895");
            newHcoElem.SetAttribute("extension", hco);

            if (hcoElement == null)
                if (ccnElement != null)
                    parent.InsertBefore(newHcoElem, ccnElement);
                else
                    parent.PrependChild(newHcoElem);
            else
            {
                parent.InsertBefore(newHcoElem, hcoElement);
                parent.RemoveChild(hcoElement);
            }

        }

        private void FormatMeasureSection(XmlDocument doc, List<Guid> measureList, XmlNamespaceManager nsmgr)
        {

            XmlElement measureSection = (XmlElement)doc.SelectSingleNode(
                ConvertToXPath("ClinicalDocument/component/structuredBody/component/section/templateId[@root = '2.16.840.1.113883.10.20.24.2.2']", _defaultPrefix),
                nsmgr).ParentNode;

            var tjcTemplate = measureSection.SelectSingleNode(ConvertToXPath("templateId[@root = '1.3.6.1.4.1.33895.1.6']", _defaultPrefix), nsmgr);
            XmlElement newTjcTemplate = doc.CreateElement("templateId", nsmgr.LookupNamespace(_defaultPrefix));
            newTjcTemplate.SetAttribute("root", "1.3.6.1.4.1.33895.1.6");

            if (tjcTemplate == null)
                measureSection.PrependChild(newTjcTemplate);
            else
            {
                measureSection.InsertAfter(newTjcTemplate, tjcTemplate);
                measureSection.RemoveChild(tjcTemplate);
            }

            var path = ConvertToXPath("text/table/tbody/tr", _defaultPrefix);
            XmlNodeList nodeList = measureSection.SelectNodes(path, nsmgr);
            foreach (XmlElement item in nodeList)
            {
                var value = item.LastChild.InnerText;

                Guid measureGuid;
                if (Guid.TryParse(value, out measureGuid))
                {
                    if (!measureList.Contains(measureGuid))
                    {
                        item.ParentNode.RemoveChild(item);
                    }
                }
            }

            var entryPath = ConvertToXPath("entry/organizer/reference/externalDocument/id[@root = '2.16.840.1.113883.4.738']", _defaultPrefix);
            var entryNodeList = measureSection.SelectNodes(entryPath, nsmgr);
            foreach (XmlElement item in entryNodeList)
            {
                string value = item.GetAttribute("extension");

                Guid measureGuid;
                if (Guid.TryParse(value, out measureGuid))
                {
                    if (!measureList.Contains(measureGuid))
                    {
                        XmlElement entry = (XmlElement)item.ParentNode.ParentNode.ParentNode.ParentNode;

                        //string tmpPath = ConvertToXPath(String.Format("entry[default:organizer/reference/externalDocument/id[@extension = \'{0}\']]", measureGuid), _defaultPrefix);
                        //XmlElement entry = (XmlElement)measureSection.SelectSingleNode(tmpPath, nsmgr);
                        entry.ParentNode.RemoveChild(entry);
                    }
                }
            }
        }

        private void FormatEncounterSection(XmlDocument doc, DateTime start, DateTime end, XmlNamespaceManager nsmgr)
        {
            XmlElement patientData = (XmlElement)doc.SelectSingleNode(
                ConvertToXPath("ClinicalDocument/component/structuredBody/component/section/templateId[@root = '2.16.840.1.113883.10.20.17.2.4']", _defaultPrefix),
                nsmgr).ParentNode;

            XmlNodeList nodeList = patientData.SelectNodes(ConvertToXPath("entry/encounter", _defaultPrefix), nsmgr);
            foreach (XmlElement node in nodeList)
            {
                var templateId = node.SelectSingleNode(ConvertToXPath("templateId[@root = '2.16.840.1.113883.10.20.24.3.23']", _defaultPrefix), nsmgr);
                var code = node.SelectSingleNode("default:code[@sdtc:valueSet = '2.16.840.1.113883.3.117.1.7.1.424' or @sdtc:valueSet = '2.16.840.1.113883.3.666.5.307']", nsmgr);

                if (templateId == null || code == null)
                    continue;

                var high = (XmlElement)node.SelectSingleNode(ConvertToXPath("effectiveTime/high", _defaultPrefix), nsmgr);
                var value = high.GetAttribute("value");

                if (!String.IsNullOrEmpty(value))
                {
                    var date = ParseDatetime(value);
                    if (date < start || date > end)
                        node.ParentNode.RemoveChild(node);
                }
            }
        }
        #endregion

        public void ProduceUpdatedTjcQrda1(XmlDocument originalFile, Stream outputStream, Guid vendorTrackingId, string hcoId,
            IEnumerable<Guid> selectedVersionSpecificMeasureGuids, DateTime reportingPeriodStart, DateTime reportingPeriodEnd)
        {
            XmlNamespaceManager nsmgr = SetUpNamespaces(originalFile);

            AddTjcHeader(originalFile, nsmgr);

            AddVenderTrackingId(originalFile, nsmgr, vendorTrackingId);

            RemovePhiData(originalFile, nsmgr);

            AddHco(originalFile, hcoId, nsmgr);

            FormatMeasureSection(originalFile, selectedVersionSpecificMeasureGuids.ToList(), nsmgr);

            FormatEncounterSection(originalFile, reportingPeriodStart, reportingPeriodEnd, nsmgr);

            originalFile.Save(outputStream);
        }
    }
}
