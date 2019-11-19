using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Xml;

namespace csharprdl
{
   class RdlGenerator
   {
      SqlConnection m_connection;
      string m_connectString;
      string m_commandText;
      ArrayList m_fields;

      public static void Main()
      {
         RdlGenerator myRdlGenerator = new RdlGenerator();
         myRdlGenerator.Run();
      }

      public void Run()
      {
         try
         {
            // Call methods to create the RDL
            this.OpenConnection();
            this.GenerateFieldsList();
            this.GenerateRdl();

            Console.WriteLine("RDL file generated successfully.");
         }

         catch (Exception exception)
         {
            Console.WriteLine("An error occurred: " + exception.Message);
         }

         finally
         {
            // Close the connection string
            m_connection.Close();
         }
      }

      public void OpenConnection()
      {
         // Create a connection object
         m_connection = new SqlConnection();
               
         // Create the connection string
         string server = "BSACCOUNTING-PC\\SQLEXPRESS"; //localhost does not work
         string database = "AdventureWorks2008R2";
         string loc = string.Format(@"data source={0};initial catalog={1};integrated security=SSPI",
            server,
            database
         );
         m_connectString = loc;
         m_connection.ConnectionString = m_connectString; 
               
         // Open the connection
         m_connection.Open();
      }

      public void GenerateFieldsList()
      {
         SqlCommand command;
         SqlDataReader reader;

         // Executing a query to retrieve a fields list for the report
         command = m_connection.CreateCommand();
         m_commandText = 
            "SELECT Person.CountryRegion.Name AS CountryName, Person.StateProvince.Name AS StateProvince " +
            "FROM Person.StateProvince " +
            "INNER JOIN Person.CountryRegion ON Person.StateProvince.CountryRegionCode = Person.CountryRegion.CountryRegionCode " +
            "ORDER BY Person.CountryRegion.Name";
         command.CommandText = m_commandText;
         
         // Execute and create a reader for the current command
         reader = command.ExecuteReader(CommandBehavior.SchemaOnly);
         
         // For each field in the resultset, add the name to an array list
         m_fields = new ArrayList();
         for (int i = 0; i <= reader.FieldCount - 1; i++)
         {
            m_fields.Add(reader.GetName(i));
         }
      }

      public void GenerateRdl()
      {
         // Create an XML document
         XmlDocument doc = new XmlDocument();
         string xmlData = "<Report " +
         "xmlns=\"http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition\">" +
            "</Report>";
         doc.Load(new StringReader(xmlData));

         // Report element
         XmlElement report = (XmlElement)doc.FirstChild;
         AddElement(report, "AutoRefresh", "0");
         AddElement(report, "ConsumeContainerWhitespace", "true");

         //DataSources element
         XmlElement dataSources = AddElement(report, "DataSources", null);
         //DataSource element
         XmlElement dataSource = AddElement(dataSources, "DataSource", null);
         XmlAttribute attr = dataSource.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "DataSource1";
         XmlElement connectionProperties = AddElement(dataSource, "ConnectionProperties", null);
         AddElement(connectionProperties, "DataProvider", "SQL");
         AddElement(connectionProperties, "ConnectString", m_connectString);
         AddElement(connectionProperties, "IntegratedSecurity", "true");
         //DataSets element
         XmlElement dataSets = AddElement(report, "DataSets", null);
         XmlElement dataSet = AddElement(dataSets, "DataSet", null);
         attr = dataSet.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "DataSet1";
         //Query element
         XmlElement query = AddElement(dataSet, "Query", null);
         AddElement(query, "DataSourceName", "DataSource1");
         AddElement(query, "CommandText", m_commandText);
         AddElement(query, "Timeout", "30");
         //Fields element
         XmlElement fields = AddElement(dataSet, "Fields", null);
         XmlElement field = AddElement(fields, "Field", null);
         attr = field.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "CountryName";
         AddElement(field, "DataField", "CountryName");
         field = AddElement(fields, "Field", null);
         attr = field.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "StateProvince";
         AddElement(field, "DataField", "StateProvince");

         //end of DataSources

         //ReportSections element
         XmlElement reportSections = AddElement(report, "ReportSections", null);
         XmlElement reportSection = AddElement(reportSections, "ReportSection", null);
         AddElement(reportSection, "Width", "6in");
         AddElement(reportSection, "Page", null);
         XmlElement body = AddElement(reportSection, "Body", null);
         AddElement(body, "Height", "1.5in");
         XmlElement reportItems = AddElement(body, "ReportItems", null);
         // Tablix element
         XmlElement tablix = AddElement(reportItems, "Tablix", null);
         attr = tablix.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "Tablix1";
         AddElement(tablix, "DataSetName", "DataSet1");
         AddElement(tablix, "Top", "0.5in");
         AddElement(tablix, "Left", "0.5in");
         AddElement(tablix, "Height", "0.5in");
         AddElement(tablix, "Width", "3in");
         XmlElement tablixBody = AddElement(tablix, "TablixBody", null);
         //TablixColumns element
         XmlElement tablixColumns = AddElement(tablixBody, "TablixColumns", null);
         XmlElement tablixColumn = AddElement(tablixColumns, "TablixColumn", null);
         AddElement(tablixColumn, "Width", "1.5in");
         tablixColumn = AddElement(tablixColumns, "TablixColumn", null);
         AddElement(tablixColumn, "Width", "1.5in");
         //TablixRows element
         XmlElement tablixRows = AddElement(tablixBody, "TablixRows", null);
         //TablixRow element (header row)
         XmlElement tablixRow = AddElement(tablixRows, "TablixRow", null);
         AddElement(tablixRow, "Height", "0.5in");
         XmlElement tablixCells = AddElement(tablixRow, "TablixCells", null);
         // TablixCell element (first cell)
         XmlElement tablixCell = AddElement(tablixCells, "TablixCell", null);
         XmlElement cellContents = AddElement(tablixCell, "CellContents", null);
         XmlElement textbox = AddElement(cellContents, "Textbox", null);
         attr = textbox.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "HeaderCountryName";
         AddElement(textbox, "KeepTogether", "true");
         XmlElement paragraphs = AddElement(textbox, "Paragraphs", null);
         XmlElement paragraph = AddElement(paragraphs, "Paragraph", null);
         XmlElement textRuns = AddElement(paragraph, "TextRuns", null);
         XmlElement textRun = AddElement(textRuns, "TextRun", null);
         AddElement(textRun, "Value", "CountryName");
         XmlElement style = AddElement(textRun, "Style", null);
         AddElement(style, "TextDecoration", "Underline");
         // TablixCell element (second cell)
         tablixCell = AddElement(tablixCells, "TablixCell", null);
         cellContents = AddElement(tablixCell, "CellContents", null);
         textbox = AddElement(cellContents, "Textbox", null);
         attr = textbox.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "HeaderStateProvince";
         AddElement(textbox, "KeepTogether", "true");
         paragraphs = AddElement(textbox, "Paragraphs", null);
         paragraph = AddElement(paragraphs, "Paragraph", null);
         textRuns = AddElement(paragraph, "TextRuns", null);
         textRun = AddElement(textRuns, "TextRun", null);
         AddElement(textRun, "Value", "StateProvince");
         style = AddElement(textRun, "Style", null);
         AddElement(style, "TextDecoration", "Underline");
         //TablixRow element (details row)
         tablixRow = AddElement(tablixRows, "TablixRow", null);
         AddElement(tablixRow, "Height", "0.5in");
         tablixCells = AddElement(tablixRow, "TablixCells", null);
         // TablixCell element (first cell)
         tablixCell = AddElement(tablixCells, "TablixCell", null);
         cellContents = AddElement(tablixCell, "CellContents", null);
         textbox = AddElement(cellContents, "Textbox", null);
         attr = textbox.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "CountryName";
         AddElement(textbox, "HideDuplicates", "DataSet1");
         AddElement(textbox, "KeepTogether", "true");
         paragraphs = AddElement(textbox, "Paragraphs", null);
         paragraph = AddElement(paragraphs, "Paragraph", null);
         textRuns = AddElement(paragraph, "TextRuns", null);
         textRun = AddElement(textRuns, "TextRun", null);
         AddElement(textRun, "Value", "=Fields!CountryName.Value");
         style = AddElement(textRun, "Style", null);
         // TablixCell element (second cell)
         tablixCell = AddElement(tablixCells, "TablixCell", null);
         cellContents = AddElement(tablixCell, "CellContents", null);
         textbox = AddElement(cellContents, "Textbox", null);
         attr = textbox.Attributes.Append(doc.CreateAttribute("Name"));
         attr.Value = "StateProvince";
         AddElement(textbox, "HideDuplicates", "DataSet1");
         AddElement(textbox, "KeepTogether", "true");
         paragraphs = AddElement(textbox, "Paragraphs", null);
         paragraph = AddElement(paragraphs, "Paragraph", null);
         textRuns = AddElement(paragraph, "TextRuns", null);
         textRun = AddElement(textRuns, "TextRun", null);
         AddElement(textRun, "Value", "=Fields!StateProvince.Value");
         style = AddElement(textRun, "Style", null);
         //End of second row

         //End of TablixBody

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

         //End of Tablix, ReportItems, ReportSections

         //Save XML document to file
         doc.Save("Report1.rdl");
      }

      public XmlElement AddElement(XmlElement parent, string name, string value)
      {
         XmlElement newelement = parent.OwnerDocument.CreateElement(name,
            "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition");
         parent.AppendChild(newelement);
         if (value != null) newelement.InnerText = value;
         return newelement;
      }
   }
}