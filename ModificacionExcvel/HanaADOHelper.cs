using Sap.Data.Hana;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModificacionExcvel
{
    public delegate T GetValues<T>(Dictionary<string, string> prm);

    public class HanaADOHelper
    {
        public IEnumerable<T> GetResultAsType<T>(string qry, GetValues<T> getValues, params string[] prms)
        {
            HanaConnectionManager hcm = new HanaConnectionManager();
            HanaConnection hc = null;

            try
            {
                hc = hcm.GetConnection();
                HanaCommand cmd = new HanaCommand(GetSqlQry(qry, prms), hc);
                hcm.OpenConnection();
                HanaDataReader hdr = cmd.ExecuteReader();
                //List<T> lstTemp = new List<T>();
                while (hdr.Read())
                {
                    var dc = new Dictionary<string, string>();

                    for (int i = 0; i < hdr.FieldCount; i++)
                    {
                        try
                        {
                            // Verifica si el valor del campo es DBNull.Value antes de convertirlo a cadena
                            object fieldValue = hdr[i];
                            dc[hdr.GetName(i)] = fieldValue != DBNull.Value ? fieldValue.ToString() : "";
                        }
                        catch (Exception)
                        {

                        }
                    }

                    yield return getValues(dc);
                }
            }
            finally
            {
                if (hc?.State == System.Data.ConnectionState.Open)
                    hcm.CloseConnection();
            }

        }

        public string insertValueSql(string qry, params object[] prms)
        {
            HanaConnectionManager hcm = new HanaConnectionManager();
            HanaConnection hc = null;
            string result = null;
            try
            {
                hc = hcm.GetConnection();
                HanaCommand cmd = new HanaCommand(GetSqlQry(qry, prms), hc);
                hcm.OpenConnection();
                HanaDataReader hdr = cmd.ExecuteReader();
                while (hdr.Read())
                {
                    result = hdr.GetString(0);
                }
                return result;
            }
            finally
            {
                if (hc?.State == System.Data.ConnectionState.Open)
                    hcm.CloseConnection();
            }
        }

        private string GetSqlQry(string sqlCommand, object[] prms)
        {
            for (int i = 0; i < prms.Length; i++)
            {
                // Reemplazar valores nulos con un marcador especial
                prms[i] = prms[i] ?? "NULL_MARKER";
            }

            // Formatear la cadena de SQL
            string formattedSql = string.Format(sqlCommand, prms);

            // Reemplazar los marcadores especiales con NULL
            formattedSql = formattedSql.Replace("'NULL_MARKER'", "NULL");

            return formattedSql;
        }

        public string GetValueSql(string qry, params string[] prms)
        {
            HanaConnectionManager hcm = new HanaConnectionManager();
            HanaConnection hc = null;
            string result = null;
            try
            {
                hc = hcm.GetConnection();
                HanaCommand cmd = new HanaCommand(GetSqlQry(qry, prms), hc);
                hcm.OpenConnection();
                HanaDataReader hdr = cmd.ExecuteReader();
                while (hdr.Read())
                {
                    result = hdr.GetString(0);
                }
                return result;
            }
            finally
            {
                if (hc?.State == System.Data.ConnectionState.Open)
                    hcm.CloseConnection();
            }
        }

        private string GetSqlQry(string sqlCommand, string[] prms)
        {
            return string.Format(sqlCommand, prms);
        }
    }
}
