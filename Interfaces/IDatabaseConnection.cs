using System.Collections;

namespace RDLApp.Interface{
    public interface IDatabaseConnection{
        public void set_connection_string(string conn);
        public ArrayList exceute_command_schema_only(string command);
        public string get_connection_string();
        public string get_command_string();
        public void close_connection();
    }
}