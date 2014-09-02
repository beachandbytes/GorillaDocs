using GorillaDocs.Word;
using NUnit.Framework;
using Wd = Microsoft.Office.Interop.Word;
using System;
using System.Security;
using System.Threading;
using System.Xml.Linq;
using System.Linq;

namespace GorillaDocs.IntegrationTests.Word
{
    [TestFixture]
    public class SharePointSaveTests
    {
        const string template = @"C:\Repos\MacroView\VS2010\BMGlobal.Office\BMGlobal.Common\Office Folders\Templates\Global\Correspondence\Blank.dotm";
        const string library = "http://mvperthsp2010.cloudapp.net/Documents/";
        const string contentTypeId = "0x01010045DBC7812A60A44F9523802FF0F461DC";
        Wd.Application wordApp;

        [SetUp]
        public void setup()
        {
            wordApp = WordApplicationHelper.GetWordApplication();
            Thread.Sleep(3000);
        }

        [Test]
        /// <summary>
        /// This causes a DIP error
        /// </summary>
        public void Save_with_ContentType_docprop_and_full_schema()
        {
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetSchema);

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            dlg.Show();

            wordApp.DisplayDocumentInformationPanel = true;
        }

        [Test]
        /// <summary>
        /// DIP works well but ignores default value for Enterprise Keywords
        /// </summary>
        public void Save_with_ContentType_docprop_and_full_schema_with_reload()
        {
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetSchema);

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            if (dlg.Show() == -1)
                doc.Reload();

            wordApp.DisplayDocumentInformationPanel = true;
        }

        [Test]
        /// <summary>
        /// DIP works well but ignores default value for Enterprise Keywords
        /// </summary>
        public void Save_with_ContentType_docprop_and_minimal_schema()
        {
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetMinimalSchema(contentTypeId));

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            dlg.Show();

            wordApp.DisplayDocumentInformationPanel = true;
        }

        /// <summary>
        /// DIP works well, but no default values
        /// </summary>
        [Test]
        public void Save_with_Fax_set()
        {
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetMinimalSchema(contentTypeId));
            doc.CustomXMLParts.Add(GetDocumentManagementSchema(GetColumnXml("bm_DocumentType", GetFieldNamespace("bm_DocumentType"), "Fax")));

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            dlg.Show();
            wordApp.DisplayDocumentInformationPanel = true;
        }

        /// <summary>
        /// DIP works well and default values are loaded in through the reload
        /// </summary>
        [Test]
        public void Save_with_Fax_set_and_reload()
        {
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetMinimalSchema(contentTypeId));
            doc.CustomXMLParts.Add(GetDocumentManagementSchema(GetColumnXml("bm_DocumentType", GetFieldNamespace("bm_DocumentType"), "Fax")));

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            dlg.Show();
            doc.Reload();

            wordApp.DisplayDocumentInformationPanel = true;
        }

        /// <summary>
        /// DIP works well and default values are loaded in through the reopen
        /// </summary>
        [Test]
        public void Save_with_Fax_set_then_close_and_reopen()
        {
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetMinimalSchema(contentTypeId));
            doc.CustomXMLParts.Add(GetDocumentManagementSchema(GetColumnXml("bm_DocumentType", GetFieldNamespace("bm_DocumentType"), "Fax")));

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            dlg.Show();

            var fullname = doc.FullName;
            doc.Saved = true;
            doc.Close(Wd.WdSaveOptions.wdDoNotSaveChanges);
            wordApp.Documents.Open(fullname);
            
            wordApp.DisplayDocumentInformationPanel = true;
        }

        /// <summary>
        /// DIP Works well
        /// </summary>
        [Test]
        public void Save_with_all_Document_Management_columns()
        {
            string xml = string.Format(@"<ifb06c1d85db408d87c509757cbefe0b xmlns='4c2493cf-19d2-4ec5-812f-ebc448b57bfc'>
                                <Terms xmlns='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'></Terms>
                            </ifb06c1d85db408d87c509757cbefe0b>
                            <bm_DocumentType xmlns='{0}' xsi:nil='true'/>
                            <TaxKeywordTaxHTField xmlns='cdfba795-2775-412f-8e1b-81ca13f854d4'>
                                <Terms xmlns='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'></Terms>
                            </TaxKeywordTaxHTField>
                            <TaxCatchAll xmlns='cdfba795-2775-412f-8e1b-81ca13f854d4'/>",GetFieldNamespace("bm_DocumentType"));
            
            var doc = wordApp.Documents.Add(template);
            doc.SetDocProp("ContentTypeId", contentTypeId);
            doc.CustomXMLParts.Add(GetMinimalSchema(contentTypeId));
            doc.CustomXMLParts.Add(GetDocumentManagementSchema(xml));

            Wd.Dialog dlg = doc.Application.Dialogs[Wd.WdWordDialog.wdDialogFileSaveAs];
            dlg.SetName(library);
            dlg.Show();

            wordApp.DisplayDocumentInformationPanel = true;
        }

        #region Helpers

        //TODO: Move to GorillaDocs assembly

        static string GetFieldNamespace(string field)
        {
            var doc = XDocument.Parse(GetSchema);
            var complexType = doc.Descendants().FirstOrDefault(p => p.Attribute("ref") != null && ((string)p.Attribute("ref")).EndsWith(field));
            var attribute = complexType.Attribute("ref");
            var prefix = attribute.Value.Substring(0, attribute.Value.IndexOf(":"));
            var schema = doc.Descendants().SingleOrDefault(p => p.Attribute("targetNamespace") != null && 
                ((string)p.Attribute("targetNamespace")) == "http://schemas.microsoft.com/office/2006/metadata/properties");
            return schema.Attribute(XNamespace.Xmlns + prefix).Value;
        }

        static string GetSchema
        {
            get
            {
                return @"<?xml version='1.0' encoding='utf-8'?><ct:contentTypeSchema ct:_='' ma:_='' ma:contentTypeName='Document' ma:contentTypeID='0x01010045DBC7812A60A44F9523802FF0F461DC' ma:contentTypeVersion='9' ma:contentTypeDescription='Create a new document.' ma:contentTypeScope='' ma:versionID='62589bf4e391605a1699ae705a3b331b' xmlns:ct='http://schemas.microsoft.com/office/2006/metadata/contentType' xmlns:ma='http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes'>
<xsd:schema targetNamespace='http://schemas.microsoft.com/office/2006/metadata/properties' ma:root='true' ma:fieldsID='1a2be3caaa6539a0f25d44e413f2bbfd' ns1:_='' ns2:_='' ns3:_='' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:p='http://schemas.microsoft.com/office/2006/metadata/properties' xmlns:ns1='http://schemas.microsoft.com/sharepoint/v3' xmlns:ns2='cdfba795-2775-412f-8e1b-81ca13f854d4' xmlns:ns3='4c2493cf-19d2-4ec5-812f-ebc448b57bfc'>
<xsd:import namespace='http://schemas.microsoft.com/sharepoint/v3'/>
<xsd:import namespace='cdfba795-2775-412f-8e1b-81ca13f854d4'/>
<xsd:import namespace='4c2493cf-19d2-4ec5-812f-ebc448b57bfc'/>
<xsd:element name='properties'>
<xsd:complexType>
<xsd:sequence>
<xsd:element name='documentManagement'>
<xsd:complexType>
<xsd:all>
<xsd:element ref='ns2:_dlc_DocId' minOccurs='0'/>
<xsd:element ref='ns2:_dlc_DocIdUrl' minOccurs='0'/>
<xsd:element ref='ns2:_dlc_DocIdPersistId' minOccurs='0'/>
<xsd:element ref='ns1:AverageRating' minOccurs='0'/>
<xsd:element ref='ns1:RatingCount' minOccurs='0'/>
<xsd:element ref='ns2:TaxKeywordTaxHTField' minOccurs='0'/>
<xsd:element ref='ns2:TaxCatchAll' minOccurs='0'/>
<xsd:element ref='ns3:ifb06c1d85db408d87c509757cbefe0b' minOccurs='0'/>
<xsd:element ref='ns3:bm_DocumentType' minOccurs='0'/>
</xsd:all>
</xsd:complexType>
</xsd:element>
</xsd:sequence>
</xsd:complexType>
</xsd:element>
</xsd:schema>
<xsd:schema targetNamespace='http://schemas.microsoft.com/sharepoint/v3' elementFormDefault='qualified' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:dms='http://schemas.microsoft.com/office/2006/documentManagement/types' xmlns:pc='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'>
<xsd:import namespace='http://schemas.microsoft.com/office/2006/documentManagement/types'/>
<xsd:import namespace='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'/>
<xsd:element name='AverageRating' ma:index='11' nillable='true' ma:displayName='Rating (0-5)' ma:decimals='2' ma:description='Average value of all the ratings that have been submitted' ma:indexed='true' ma:internalName='AverageRating' ma:readOnly='true'>
<xsd:simpleType>
<xsd:restriction base='dms:Number'/>
</xsd:simpleType>
</xsd:element>
<xsd:element name='RatingCount' ma:index='12' nillable='true' ma:displayName='Number of Ratings' ma:decimals='0' ma:description='Number of ratings submitted' ma:internalName='RatingCount' ma:readOnly='true'>
<xsd:simpleType>
<xsd:restriction base='dms:Number'/>
</xsd:simpleType>
</xsd:element>
</xsd:schema>
<xsd:schema targetNamespace='cdfba795-2775-412f-8e1b-81ca13f854d4' elementFormDefault='qualified' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:dms='http://schemas.microsoft.com/office/2006/documentManagement/types' xmlns:pc='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'>
<xsd:import namespace='http://schemas.microsoft.com/office/2006/documentManagement/types'/>
<xsd:import namespace='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'/>
<xsd:element name='_dlc_DocId' ma:index='8' nillable='true' ma:displayName='Document ID Value' ma:description='The value of the document ID assigned to this item.' ma:internalName='_dlc_DocId' ma:readOnly='true'>
<xsd:simpleType>
<xsd:restriction base='dms:Text'/>
</xsd:simpleType>
</xsd:element>
<xsd:element name='_dlc_DocIdUrl' ma:index='9' nillable='true' ma:displayName='Document ID' ma:description='Permanent link to this document.' ma:hidden='true' ma:internalName='_dlc_DocIdUrl' ma:readOnly='true'>
<xsd:complexType>
<xsd:complexContent>
<xsd:extension base='dms:URL'>
<xsd:sequence>
<xsd:element name='Url' type='dms:ValidUrl' minOccurs='0' nillable='true'/>
<xsd:element name='Description' type='xsd:string' nillable='true'/>
</xsd:sequence>
</xsd:extension>
</xsd:complexContent>
</xsd:complexType>
</xsd:element>
<xsd:element name='_dlc_DocIdPersistId' ma:index='10' nillable='true' ma:displayName='Persist ID' ma:description='Keep ID on add.' ma:hidden='true' ma:internalName='_dlc_DocIdPersistId' ma:readOnly='true'>
<xsd:simpleType>
<xsd:restriction base='dms:Boolean'/>
</xsd:simpleType>
</xsd:element>
<xsd:element name='TaxKeywordTaxHTField' ma:index='14' nillable='true' ma:taxonomy='true' ma:internalName='TaxKeywordTaxHTField' ma:taxonomyFieldName='TaxKeyword' ma:displayName='Enterprise Keywords' ma:fieldId='{23f27201-bee3-471e-b2e7-b64fd8b7ca38}' ma:taxonomyMulti='true' ma:sspId='72077a88-8b3c-401d-9470-298ae828adf4' ma:termSetId='00000000-0000-0000-0000-000000000000' ma:anchorId='00000000-0000-0000-0000-000000000000' ma:open='true' ma:isKeyword='true'>
<xsd:complexType>
<xsd:sequence>
<xsd:element ref='pc:Terms' minOccurs='0' maxOccurs='1'></xsd:element>
</xsd:sequence>
</xsd:complexType>
</xsd:element>
<xsd:element name='TaxCatchAll' ma:index='15' nillable='true' ma:displayName='Taxonomy Catch All Column' ma:hidden='true' ma:list='{e9c42a04-4c83-42c7-bc4d-e33a16aa14ef}' ma:internalName='TaxCatchAll' ma:showField='CatchAllData' ma:web='cdfba795-2775-412f-8e1b-81ca13f854d4'>
<xsd:complexType>
<xsd:complexContent>
<xsd:extension base='dms:MultiChoiceLookup'>
<xsd:sequence>
<xsd:element name='Value' type='dms:Lookup' maxOccurs='unbounded' minOccurs='0' nillable='true'/>
</xsd:sequence>
</xsd:extension>
</xsd:complexContent>
</xsd:complexType>
</xsd:element>
</xsd:schema>
<xsd:schema targetNamespace='4c2493cf-19d2-4ec5-812f-ebc448b57bfc' elementFormDefault='qualified' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:dms='http://schemas.microsoft.com/office/2006/documentManagement/types' xmlns:pc='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'>
<xsd:import namespace='http://schemas.microsoft.com/office/2006/documentManagement/types'/>
<xsd:import namespace='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'/>
<xsd:element name='ifb06c1d85db408d87c509757cbefe0b' ma:index='17' nillable='true' ma:taxonomy='true' ma:internalName='ifb06c1d85db408d87c509757cbefe0b' ma:taxonomyFieldName='Country' ma:displayName='Country' ma:default='' ma:fieldId='{2fb06c1d-85db-408d-87c5-09757cbefe0b}' ma:taxonomyMulti='true' ma:sspId='72077a88-8b3c-401d-9470-298ae828adf4' ma:termSetId='d4dc0692-dda2-40eb-bdad-6f7e69103556' ma:anchorId='00000000-0000-0000-0000-000000000000' ma:open='false' ma:isKeyword='false'>
<xsd:complexType>
<xsd:sequence>
<xsd:element ref='pc:Terms' minOccurs='0' maxOccurs='1'></xsd:element>
</xsd:sequence>
</xsd:complexType>
</xsd:element>
<xsd:element name='bm_DocumentType' ma:index='18' nillable='true' ma:displayName='Document Type' ma:format='Dropdown' ma:internalName='bm_DocumentType'>
<xsd:simpleType>
<xsd:restriction base='dms:Choice'>
<xsd:enumeration value='Letter'/>
<xsd:enumeration value='Fax'/>
<xsd:enumeration value='Memo'/>
<xsd:enumeration value='Blank'/>
</xsd:restriction>
</xsd:simpleType>
</xsd:element>
</xsd:schema>
<xsd:schema targetNamespace='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' elementFormDefault='qualified' attributeFormDefault='unqualified' blockDefault='#all' xmlns='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/' xmlns:odoc='http://schemas.microsoft.com/internal/obd'>
<xsd:import namespace='http://purl.org/dc/elements/1.1/' schemaLocation='http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd'/>
<xsd:import namespace='http://purl.org/dc/terms/' schemaLocation='http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcterms.xsd'/>
<xsd:element name='coreProperties' type='CT_coreProperties'/>
<xsd:complexType name='CT_coreProperties'>
<xsd:all>
<xsd:element ref='dc:creator' minOccurs='0' maxOccurs='1'/>
<xsd:element ref='dcterms:created' minOccurs='0' maxOccurs='1'/>
<xsd:element ref='dc:identifier' minOccurs='0' maxOccurs='1'/>
<xsd:element name='contentType' minOccurs='0' maxOccurs='1' type='xsd:string' ma:index='0' ma:displayName='Content Type'/>
<xsd:element ref='dc:title' minOccurs='0' maxOccurs='1' ma:index='4' ma:displayName='Title'/>
<xsd:element ref='dc:subject' minOccurs='0' maxOccurs='1'/>
<xsd:element ref='dc:description' minOccurs='0' maxOccurs='1'/>
<xsd:element name='keywords' minOccurs='0' maxOccurs='1' type='xsd:string'/>
<xsd:element ref='dc:language' minOccurs='0' maxOccurs='1'/>
<xsd:element name='category' minOccurs='0' maxOccurs='1' type='xsd:string'/>
<xsd:element name='version' minOccurs='0' maxOccurs='1' type='xsd:string'/>
<xsd:element name='revision' minOccurs='0' maxOccurs='1' type='xsd:string'>
<xsd:annotation>
<xsd:documentation>
                        This value indicates the number of saves or revisions. The application is responsible for updating this value after each revision.
                    </xsd:documentation>
</xsd:annotation>
</xsd:element>
<xsd:element name='lastModifiedBy' minOccurs='0' maxOccurs='1' type='xsd:string'/>
<xsd:element ref='dcterms:modified' minOccurs='0' maxOccurs='1'/>
<xsd:element name='contentStatus' minOccurs='0' maxOccurs='1' type='xsd:string'/>
</xsd:all>
</xsd:complexType>
</xsd:schema>
<xs:schema targetNamespace='http://schemas.microsoft.com/office/infopath/2007/PartnerControls' elementFormDefault='qualified' attributeFormDefault='unqualified' xmlns:pc='http://schemas.microsoft.com/office/infopath/2007/PartnerControls' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
<xs:element name='Person'>
<xs:complexType>
<xs:sequence>
<xs:element ref='pc:DisplayName' minOccurs='0'></xs:element>
<xs:element ref='pc:AccountId' minOccurs='0'></xs:element>
<xs:element ref='pc:AccountType' minOccurs='0'></xs:element>
</xs:sequence>
</xs:complexType>
</xs:element>
<xs:element name='DisplayName' type='xs:string'></xs:element>
<xs:element name='AccountId' type='xs:string'></xs:element>
<xs:element name='AccountType' type='xs:string'></xs:element>
<xs:element name='BDCAssociatedEntity'>
<xs:complexType>
<xs:sequence>
<xs:element ref='pc:BDCEntity' minOccurs='0' maxOccurs='unbounded'></xs:element>
</xs:sequence>
<xs:attribute ref='pc:EntityNamespace'></xs:attribute>
<xs:attribute ref='pc:EntityName'></xs:attribute>
<xs:attribute ref='pc:SystemInstanceName'></xs:attribute>
<xs:attribute ref='pc:AssociationName'></xs:attribute>
</xs:complexType>
</xs:element>
<xs:attribute name='EntityNamespace' type='xs:string'></xs:attribute>
<xs:attribute name='EntityName' type='xs:string'></xs:attribute>
<xs:attribute name='SystemInstanceName' type='xs:string'></xs:attribute>
<xs:attribute name='AssociationName' type='xs:string'></xs:attribute>
<xs:element name='BDCEntity'>
<xs:complexType>
<xs:sequence>
<xs:element ref='pc:EntityDisplayName' minOccurs='0'></xs:element>
<xs:element ref='pc:EntityInstanceReference' minOccurs='0'></xs:element>
<xs:element ref='pc:EntityId1' minOccurs='0'></xs:element>
<xs:element ref='pc:EntityId2' minOccurs='0'></xs:element>
<xs:element ref='pc:EntityId3' minOccurs='0'></xs:element>
<xs:element ref='pc:EntityId4' minOccurs='0'></xs:element>
<xs:element ref='pc:EntityId5' minOccurs='0'></xs:element>
</xs:sequence>
</xs:complexType>
</xs:element>
<xs:element name='EntityDisplayName' type='xs:string'></xs:element>
<xs:element name='EntityInstanceReference' type='xs:string'></xs:element>
<xs:element name='EntityId1' type='xs:string'></xs:element>
<xs:element name='EntityId2' type='xs:string'></xs:element>
<xs:element name='EntityId3' type='xs:string'></xs:element>
<xs:element name='EntityId4' type='xs:string'></xs:element>
<xs:element name='EntityId5' type='xs:string'></xs:element>
<xs:element name='Terms'>
<xs:complexType>
<xs:sequence>
<xs:element ref='pc:TermInfo' minOccurs='0' maxOccurs='unbounded'></xs:element>
</xs:sequence>
</xs:complexType>
</xs:element>
<xs:element name='TermInfo'>
<xs:complexType>
<xs:sequence>
<xs:element ref='pc:TermName' minOccurs='0'></xs:element>
<xs:element ref='pc:TermId' minOccurs='0'></xs:element>
</xs:sequence>
</xs:complexType>
</xs:element>
<xs:element name='TermName' type='xs:string'></xs:element>
<xs:element name='TermId' type='xs:string'></xs:element>
</xs:schema>
</ct:contentTypeSchema>";
            }
        }

        static string GetMinimalSchema(string contentTypeId)
        {
            return String.Format(@"<?xml version='1.0' encoding='utf-8'?>
                                               <ct:contentTypeSchema 
                                                   ct:_='' 
                                                   ma:_='' 
                                                   ma:contentTypeName='' 
                                                   ma:contentTypeID='{0}' 
                                                   ma:contentTypeVersion='' 
                                                   ma:contentTypeDescription='' 
                                                   ma:contentTypeScope='' 
                                                   ma:versionID='' 
                                                   xmlns:ct='http://schemas.microsoft.com/office/2006/metadata/contentType' 
                                                   xmlns:ma='http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes'/>", contentTypeId);
        }

        static string GetDocumentManagementSchema(string xml)
        {
            return String.Format(@"<?xml version='1.0'?>
		                            <p:properties xmlns:p='http://schemas.microsoft.com/office/2006/metadata/properties' 
                                                  xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' 
                                                  xmlns:pc='http://schemas.microsoft.com/office/infopath/2007/PartnerControls'>
		                                <documentManagement>{0}</documentManagement>
		                            </p:properties>", xml);
        }

        static string GetColumnXml(string internalName, string namespaceId, string value)
        {
            return string.Format("<{0} xmlns='{1}' {3}>{2}</{0}>", internalName, namespaceId, SecurityElement.Escape(value), GetNil(value));
        }

        static object GetNil(string value)
        {
            if (string.IsNullOrEmpty(value) || value == ";#;#")
                return "xsi:nil='true'";
            else
                return string.Empty;
        }
        #endregion
    }
}
