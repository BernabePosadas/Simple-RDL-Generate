using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Xml;
using RDLApp.Interface;

namespace RDLApp
{
    public class RdlFileGenerator
    {
        private XmlDocument doc;
        private IDatabaseConnection conn;

        private ArrayList field_list;
        private XmlAttribute attr;
        public RdlFileGenerator(IDatabaseConnection conn, ArrayList field_list)
        {
            this.conn = conn; 
            this.field_list = field_list;
            this.doc = new XmlDocument();
            string xmlData = "<Report " +
            "xmlns=\"http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition\">" +
                "</Report>";
            doc.Load(new StringReader(xmlData));
        }
        public XmlDocument GenerateDocument()
        {
            XmlElement report = (XmlElement)this.doc.FirstChild;
            this.AddElement(report, "AutoRefresh", "0");
            this.AddElement(report, "ConsumeContainerWhitespace", "true");
            this.AddDataSourceElement(report);
            this.AddDataSetsElement(report);
            this.AddReprortSection(report);
            return doc;
        }
        private void AddDataSourceElement(XmlElement report)
        {
            //DataSources element
            XmlElement dataSources = this.AddElement(report, "DataSources", null);
            //DataSource element
            XmlElement dataSource = this.AddElement(dataSources, "DataSource", null);
            XmlAttribute attr = dataSource.Attributes.Append(this.doc.CreateAttribute("Name"));
            attr.Value = "DataSource1";
            XmlElement connectionProperties = AddElement(dataSource, "ConnectionProperties", null);
            AddElement(connectionProperties, "DataProvider", "SQL");
            AddElement(connectionProperties, "ConnectString", this.conn.get_connection_string());
            AddElement(connectionProperties, "IntegratedSecurity", "true");
            this.attr = attr;
        }
        private void AddDataSetsElement(XmlElement report)
        {
            //DataSets element
            XmlElement dataSets = this.AddElement(report, "DataSets", null);
            XmlElement dataSet = this.AddElement(dataSets, "DataSet", null);
            this.attr = dataSet.Attributes.Append(this.doc.CreateAttribute("Name"));
            this.attr.Value = "DataSet1";
            //Query element
            XmlElement query = this.AddElement(dataSet, "Query", null);
            this.AddElement(query, "DataSourceName", "DataSource1");
            this.AddElement(query, "CommandText", this.conn.get_command_string());
            this.AddElement(query, "Timeout", "30");
            //Fields element
            XmlElement fields = this.AddElement(dataSet, "Fields", null);
            foreach(var field_name in this.field_list){
                this.GenerateFieldElement(field_name.ToString(), fields);
            }
            
        }
        private void GenerateFieldElement(string ColumnName, XmlElement fields){
            XmlElement field = this.AddElement(fields, "Field", null);
            this.attr = field.Attributes.Append(this.doc.CreateAttribute("Name"));
            this.attr.Value = ColumnName;
            AddElement(field, "DataField", ColumnName);
        }
        private void AddReprortSection(XmlElement report)
        {
            //ReportSections element
            XmlElement reportSections = this.AddElement(report, "ReportSections", null);
            XmlElement reportSection = this.AddElement(reportSections, "ReportSection", null);
            this.AddElement(reportSection, "Width", "6in");
            this.AddElement(reportSection, "Page", null);
            XmlElement body = this.AddElement(reportSection, "Body", null);
            this.AddElement(body, "Height", "1.5in");
            XmlElement reportItems = this.AddElement(body, "ReportItems", null);
            // Tablix element
            XmlElement tablix = this.AddElement(reportItems, "Tablix", null);
            attr = tablix.Attributes.Append(this.doc.CreateAttribute("Name"));
            attr.Value = "Tablix1";
            this.AddElement(tablix, "DataSetName", "DataSet1");
            this.AddElement(tablix, "Top", "0.5in");
            this.AddElement(tablix, "Left", "0.5in");
            this.AddElement(tablix, "Height", "0.5in");
            this.AddElement(tablix, "Width", "3in");

            //tablix body 
            XmlElement tablixBody = this.GenerateBodyTablix(tablix);
            XmlElement tablixRows = this.AddElement(tablixBody, "TablixRows", null);
            XmlElement tablixCells = this.GenerateRowEntry(tablixRows);
            foreach(var field_name in this.field_list){
                this.GenerateHeaderRow(tablixCells, field_name.ToString());
            }
            tablixCells = this.GenerateRowEntry(tablixRows);
            foreach(var field_name in this.field_list){
                this.GenerateRowDetails(tablixCells, field_name.ToString());
            }

            //TablixColumnHierarchy element
            XmlElement tablixColumnHierarchy = AddElement(tablix, "TablixColumnHierarchy", null);
            XmlElement tablixMembers = AddElement(tablixColumnHierarchy, "TablixMembers", null);
            AddElement(tablixMembers, "TablixMember", null);
            AddElement(tablixMembers, "TablixMember", null);

            //TablixRowHierarchy element
            XmlElement tablixRowHierarchy = AddElement(tablix, "TablixRowHierarchy", null);
            tablixMembers = AddElement(tablixRowHierarchy, "TablixMembers", null);
            XmlElement tablixMember = AddElement(tablixMembers, "TablixMember", null);
            AddElement(tablixMember, "KeepWithGroup", "After");
            AddElement(tablixMember, "KeepTogether", "true");
            tablixMember = AddElement(tablixMembers, "TablixMember", null);
            AddElement(tablixMember, "DataElementName", "Detail_Collection");
            AddElement(tablixMember, "DataElementOutput", "Output");
            AddElement(tablixMember, "KeepTogether", "true");
            XmlElement group = AddElement(tablixMember, "Group", null);
            attr = group.Attributes.Append(doc.CreateAttribute("Name"));
            attr.Value = "Table1_Details_Group";
            AddElement(group, "DataElementName", "Detail");
            XmlElement tablixMembersNested = AddElement(tablixMember, "TablixMembers", null);
            AddElement(tablixMembersNested, "TablixMember", null);
        }
        private XmlElement GenerateBodyTablix(XmlElement tablix){
            XmlElement tablixBody = this.AddElement(tablix, "TablixBody", null);
            //TablixColumns element
            XmlElement tablixColumns = this.AddElement(tablixBody, "TablixColumns", null);
            XmlElement tablixColumn = this.AddElement(tablixColumns, "TablixColumn", null);
            this.AddElement(tablixColumn, "Width", "1.5in");
            tablixColumn = this.AddElement(tablixColumns, "TablixColumn", null);
            this.AddElement(tablixColumn, "Width", "1.5in");
            //TablixRows element
            return tablixBody;
        }
        private XmlElement GenerateRowEntry(XmlElement tablixRows)
        {
            XmlElement tablixRow = this.AddElement(tablixRows, "TablixRow", null);
            this.AddElement(tablixRow, "Height", "0.5in");
            return this.AddElement(tablixRow, "TablixCells", null);
        }
        private void GenerateHeaderRow(XmlElement tablixCells, string ColumnName){
            XmlElement textRun = this.GenerateTextRun(tablixCells, string.Format("Header{0}", ColumnName));
            this.AddElement(textRun, "Value", "CountryName");
            XmlElement style = this.AddElement(textRun, "Style", null);
            this.AddElement(style, "TextDecoration", "Underline");
        }
        private void GenerateRowDetails(XmlElement tablixCells, string ColumnName){
            XmlElement textRun = this.GenerateTextRun(tablixCells, ColumnName);
            AddElement(textRun, "Value", string.Format("=Fields!{0}.Value", ColumnName));
            XmlElement style = AddElement(textRun, "Style", null);
        }
        private XmlElement GenerateTextRun(XmlElement tablixCells, string TextboxValue){
            XmlElement tablixCell = this.AddElement(tablixCells, "TablixCell", null);
            XmlElement cellContents = this.AddElement(tablixCell, "CellContents", null);
            XmlElement textbox = this.AddElement(cellContents, "Textbox", null);
            attr = textbox.Attributes.Append(this.doc.CreateAttribute("Name"));
            attr.Value = TextboxValue;
            this.AddElement(textbox, "KeepTogether", "true");
            XmlElement paragraphs = this.AddElement(textbox, "Paragraphs", null);
            XmlElement paragraph = this.AddElement(paragraphs, "Paragraph", null);
            XmlElement textRuns = this.AddElement(paragraph, "TextRuns", null);
            return this.AddElement(textRuns, "TextRun", null);
        }
        private XmlElement AddElement(XmlElement parent, string name, string value)
        {
            XmlElement newelement = parent.OwnerDocument.CreateElement(name,
                "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition");
            parent.AppendChild(newelement);
            if (value != null) newelement.InnerText = value;
            return newelement;
        }
    }
}