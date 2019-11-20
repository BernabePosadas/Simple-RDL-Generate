using System;
using System.Collections;
using RDLApp.ConcreteImplementation;

namespace RDLApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SQLServerDatabaseConnection conn = new SQLServerDatabaseConnection();
            conn.set_connection_string(@"data source=localhost;initial catalog=AdventureWorks;integrated security=SSPI");
            ArrayList field_list = conn.exceute_command_schema_only(
               @"SELECT Person.CountryRegion.Name AS CountryName, Person.StateProvince.Name AS StateProvince 
               FROM Person.StateProvince
               INNER JOIN Person.CountryRegion ON Person.StateProvince.CountryRegionCode = Person.CountryRegion.CountryRegionCode 
               ORDER BY Person.CountryRegion.Name"
            );
            RdlGenerator generator = new RdlGenerator();
            generator.Run(conn, field_list);
        }
    }
}
