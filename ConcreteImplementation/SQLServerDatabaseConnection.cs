using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using RDLApp.Interface;

namespace RDLApp.ConcreteImplementation
{
    public class SQLServerDatabaseConnection : IDatabaseConnection
    {
        private SqlConnection m_connection;
        private string m_connectString { get; set; } = "";
        private string m_commandText { get; set; } = "";

        private bool connection_opened;
        public SQLServerDatabaseConnection()
        {
            this.m_connection = new SqlConnection();
            this.connection_opened = false;
        }
        public void set_connection_string(string conn)
        {
            this.m_connectString = conn;
            this.m_connection.ConnectionString = this.m_connectString;
            this.open_connection();
        }
        public string get_connection_string(){
            return this.m_connectString;
        }
        public string get_command_string(){
            return this.m_commandText;
        }
        public ArrayList exceute_command_schema_only(string command)
        {
            SqlCommand sqlcommand;
            SqlDataReader reader;
            ArrayList m_fields = new ArrayList();
            if (this.connection_opened)
            {
                sqlcommand = m_connection.CreateCommand();
                this.m_commandText = command;
                sqlcommand.CommandText = this.m_commandText;

                // Execute and create a reader for the current command
                reader = sqlcommand.ExecuteReader(CommandBehavior.SchemaOnly);

                // For each field in the resultset, add the name to an array list
                m_fields = new ArrayList();
                for (int i = 0; i <= reader.FieldCount - 1; i++)
                {
                    m_fields.Add(reader.GetName(i));
                }
                return m_fields;

            }
            else
            {
                throw new Exception("Connection not opened!");
            }

        }
        public void open_connection()
        {
            if (!this.connection_opened)
            {
                this.m_connection.Open();
                this.connection_opened = true;
            }
        }
        public void close_connection(){
            if(this.connection_opened){
                this.m_connection.Close();
                this.connection_opened = false;
            }
        } 
    }
}