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
    public class RdlGenerator
    {

        public void Run(IDatabaseConnection conn, ArrayList field_list)
        {
            try
            {
                RdlFileGenerator FileGenerator = new RdlFileGenerator(conn, field_list);
                XmlDocument doc = FileGenerator.GenerateDocument();
                doc.Save("Report1.rdl");
                Console.WriteLine("RDL file generated successfully.");
                Console.ReadKey();
            }

            catch (Exception exception)
            {
                Console.WriteLine("An error occurred: " + exception.Message);
                Console.ReadKey();
            }
            finally
            {
                conn.close_connection();
            }
        }
    }
}