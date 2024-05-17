using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;

namespace ModificacionExcvel
{
    public class Program
    {
        static void Main()
        {
            // Obtiene Tabla
            ObtieneDataTable("SIRE");

            // Obtiene Columnas
            ObtieneDataColumns("SIRE");

            ObtieneDataUDO("SIRE");
        }

        static void ObtieneDataColumns(string addon)
        {
            XmlDocument xmlDoc = new XmlDocument();
            // C:\Users\Sebastian\source\repos\STR_ADDONPERU_INSTALADOR\STR_ADDONPERU_INSTALADOR\Resources\Localizacion\UT.vte
            string path = $"D:\\SEBAS\\PlantillasAddonPE\\XMLS\\{addon}\\UF.xml";
            // Cargar el archivo XML con la codificación UTF-16
            using (XmlReader reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
            {
                xmlDoc.Load(reader);
            }

            XmlNodeList tableNodes = xmlDoc.SelectNodes("//BO/UserFieldsMD/row");

            List<SAPCOLUMN> listaColumns = new List<SAPCOLUMN>();
            List<UFD1> listCombos = new List<UFD1>();

            HanaADOHelper hsh = new HanaADOHelper();
            foreach (XmlNode tableNode in tableNodes)
            {
                SAPCOLUMN col = new SAPCOLUMN();
                // Obtener los valores de TableName, TableDescription y TableType
                col.Name = tableNode.SelectSingleNode("Name").InnerText;
                string _fieldId = hsh.insertValueSql("SELECT \"FieldID\" FROM CUFD WHERE \"AliasID\" = '{0}' and \"TableID\" = '{1}'", tableNode.SelectSingleNode("Name").InnerText, tableNode.SelectSingleNode("TableName").InnerText);
                // Obtiene Data de Columnas
                col.FieldID = _fieldId;
                col.Type = tableNode.SelectSingleNode("Type").InnerText;
                col.Size = tableNode.SelectSingleNode("Size").InnerText;
                col.Description = tableNode.SelectSingleNode("Description").InnerText;
                col.SubType = tableNode.SelectSingleNode("SubType").InnerText;
                col.TableName = tableNode.SelectSingleNode("TableName").InnerText;
                col.EditSize = tableNode.SelectSingleNode("EditSize").InnerText;
                col.Mandatory = tableNode.SelectSingleNode("Mandatory").InnerText;
                col.DefaultValue = tableNode.SelectSingleNode("DefaultValue") == null ? "" : tableNode.SelectSingleNode("DefaultValue").InnerText;
                listaColumns.Add(col);

                List<UFD1> udos1 = hsh.GetResultAsType("SELECT * FROM \"UFD1\" WHERE \"TableID\" = '{0}' AND \"FieldID\" = '{1}'", dc =>
                {
                    return new UFD1
                    {
                        TableID = dc["TableID"],
                        FieldID = dc["FieldID"],
                        IndexID = dc["IndexID"],
                        FldValue = dc["FldValue"],
                        Descr = dc["Descr"],
                        FldDate = dc["FldDate"]
                    };
                }, tableNode.SelectSingleNode("TableName").InnerText, _fieldId).ToList();

                if (udos1.Count > 0) listCombos.AddRange(udos1);
            }

            AddDataExcelCol(listaColumns, listCombos,addon);
        }

        static void AddDataExcelCol(List<SAPCOLUMN> lista, List<UFD1> listCombos, string addon)
        {
            // Ruta Excel
            string path = $"D:\\SEBAS\\PlantillasAddonPE\\Excels\\{addon}\\01 Plantilla de Campos.xlsm";

            Application excelApp = new Application();

            // Abrir el libro de trabajo
            Workbook workbook = excelApp.Workbooks.Open(path);

            // Obtener la hoja de trabajo
            Worksheet worksheet = (Worksheet)workbook.Sheets[1]; // Hoja 1

            // Leer los datos de la hoja de trabajo
            // Recorremos la lista de datos y los escribimos en las celdas correspondientes
            for (int i = 0; i < lista.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = lista[i].TableName;
                worksheet.Cells[i + 3, 2] = lista[i].FieldID;
                worksheet.Cells[i + 3, 3] = lista[i].Name;
                worksheet.Cells[i + 3, 4] = lista[i].Type;
                worksheet.Cells[i + 3, 5] = lista[i].Size;
                worksheet.Cells[i + 3, 6] = lista[i].Description;
                worksheet.Cells[i + 3, 7] = lista[i].SubType;
                worksheet.Cells[i + 3, 9] = lista[i].DefaultValue;
                worksheet.Cells[i + 3, 10] = lista[i].EditSize;
                worksheet.Cells[i + 3, 11] = lista[i].Mandatory;
            }

            worksheet = (Worksheet)workbook.Sheets[2]; // Hoja 1

            for (int i = 0; i < listCombos.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = listCombos[i].TableID;
                worksheet.Cells[i + 3, 2] = listCombos[i].FieldID;
                worksheet.Cells[i + 3, 3] = listCombos[i].IndexID;
                worksheet.Cells[i + 3, 4] = listCombos[i].FldValue;
                worksheet.Cells[i + 3, 5] = listCombos[i].Descr;
                worksheet.Cells[i + 3, 6] = listCombos[i].FldDate;
            }

            // Guardar los cambios
            workbook.Save();

            // Cerrar el libro de trabajo y la aplicación Excel
            workbook.Close();
            excelApp.Quit();

        }

        static void ObtieneDataTable(string addon)
        {
            XmlDocument xmlDoc = new XmlDocument();
            // C:\Users\Sebastian\source\repos\STR_ADDONPERU_INSTALADOR\STR_ADDONPERU_INSTALADOR\Resources\Localizacion\UT.vte
            string path = $"D:\\SEBAS\\PlantillasAddonPE\\XMLS\\{addon}\\UT.xml";
            // Cargar el archivo XML con la codificación UTF-16
            using (XmlReader reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
            {
                xmlDoc.Load(reader);
            }

            XmlNodeList tableNodes = xmlDoc.SelectNodes("//BO/UserTablesMD/row");

            List<SAPTABLE> listTables = new List<SAPTABLE>();

            foreach (XmlNode tableNode in tableNodes)
            {
                SAPTABLE table = new SAPTABLE();
                // Obtener los valores de TableName, TableDescription y TableType
                table.tableName = tableNode.SelectSingleNode("TableName").InnerText;
                table.tableDescription = tableNode.SelectSingleNode("TableDescription").InnerText;
                table.tableType = tableNode.SelectSingleNode("TableType").InnerText;
                listTables.Add(table);
            }

            //Console.WriteLine("s");

            AddDataExcelTable(listTables, addon);
        }

        static void AddDataExcelTable(List<SAPTABLE> lista, string addon)
        {
            // Ruta Excel
            string path = $"D:\\SEBAS\\PlantillasAddonPE\\Excels\\{addon}\\00 Plantilla de Tablas.xlsm";

            Application excelApp = new Application();

            // Abrir el libro de trabajo
            Workbook workbook = excelApp.Workbooks.Open(path);

            // Obtener la hoja de trabajo
            Worksheet worksheet = (Worksheet)workbook.Sheets[1]; // Hoja 1

            // Leer los datos de la hoja de trabajo
            // Recorremos la lista de datos y los escribimos en las celdas correspondientes
            for (int i = 0; i < lista.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = lista[i].tableName;
                worksheet.Cells[i + 3, 2] = lista[i].tableDescription;
                worksheet.Cells[i + 3, 3] = lista[i].tableType;
            }

            // Guardar los cambios
            workbook.Save();

            // Cerrar el libro de trabajo y la aplicación Excel
            workbook.Close();
            excelApp.Quit();

        }

        static void ObtieneDataUDO(string addon)
        {
            XmlDocument xmlDoc = new XmlDocument();
            // C:\Users\Sebastian\source\repos\STR_ADDONPERU_INSTALADOR\STR_ADDONPERU_INSTALADOR\Resources\Localizacion\UT.vte
            string path = $"D:\\SEBAS\\PlantillasAddonPE\\XMLS\\{addon}\\UO.xml";
            // Cargar el archivo XML con la codificación UTF-16
            using (XmlReader reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
            {
                xmlDoc.Load(reader);
            }

            XmlNodeList tableNodes = xmlDoc.SelectNodes("//BO/OUDO/row");

            List<SAPOBJECT> listaObject = new List<SAPOBJECT>();
            List<UDO1> listaudo1 = new List<UDO1>();
            List<UDO2> listaudo2 = new List<UDO2>();
            List<UDO3> listaudo3 = new List<UDO3>();

            HanaADOHelper hsh = new HanaADOHelper();

            foreach (XmlNode tableNode in tableNodes)
            {
                SAPOBJECT col = new SAPOBJECT();
                // Obtener los valores de TableName, TableDescription y TableType
                string code = tableNode.SelectSingleNode("Code").InnerText;

                SAPOBJECT ObjectEnSAP = hsh.GetResultAsType("SELECT \"Code\",\"Name\",\"TableName\",\"LogTable\",CASE WHEN \"TYPE\" = '3' " +
                    "THEN 'boud_Document' ELSE 'boud_MasterData' END AS \"TYPE\",\r\n\"MngSeries\",\"CanDelete\",\"CanClose\",\"CanCancel\"," +
                    "\"ExtName\",\"CanFind\",\"CanYrTrnsf\",\"CanDefForm\",\"CanLog\",\r\n\"OvrWrtDll\",\"UIDFormat\",\"CanArchive\"," +
                    "\"MenuItem\",\"MenuCapt\",\"FatherMenu\",\"Position\",\"CanNewForm\",\r\n\"IsRebuild\",\"NewFormSrf\",\"MenuUid\" " +
                    "FROM OUDO WHERE \"Code\" = '{0}'",
                    dc => 
                    {
                        return new SAPOBJECT 
                        {
                            Code = dc["Code"],
                            Name = dc["Name"],
                            TableName = dc["TableName"],
                            LogTable = dc["LogTable"],
                            TYPE = dc["TYPE"],
                            MngSeries = dataUdo(dc["MngSeries"]),
                            CanDelete = dataUdo(dc["CanDelete"]),
                            CanClose = dataUdo(dc["CanClose"]),
                            CanCancel = dataUdo(dc["CanCancel"]),
                            ExtName = dc["ExtName"],
                            CanFind = dataUdo(dc["CanFind"]),
                            CanYrTrnsf = dataUdo(dc["CanYrTrnsf"]),
                            CanDefForm = dataUdo(dc["CanDefForm"]),
                            CanLog = dataUdo(dc["CanLog"]),
                            OvrWrtDll = dataUdo(dc["OvrWrtDll"]),
                            UIDFormat = dataUdo(dc["UIDFormat"]),
                            CanArchive = dataUdo(dc["CanArchive"]),
                            MenuItem = dataUdo(dc["MenuItem"]),
                            MenuCapt = dc["MenuCapt"],
                            FatherMenu = dc["FatherMenu"],
                            Position = dc["Position"],
                            CanNewForm = dataUdo(dc["CanNewForm"]),
                            IsRebuild = dataUdo(dc["IsRebuild"]),
                            NewFormSrf = dc["NewFormSrf"],
                            MenuUid = dc["MenuUid"]
                        };
                    }, code).ToList()[0];

                List<UDO1> udos1 = hsh.GetResultAsType("SELECT * FROM UDO1 WHERE \"Code\" = '{0}'", dc =>
                {
                    return new UDO1
                    {
                        Code = dc["Code"],
                        LineNum = dc["SonNum"],
                        TableName = dc["TableName"],
                        LogName = dc["LogName"],
                        SonName = dc["SonName"]
                    };
                }, code).ToList();
                List<UDO2> udos2 = hsh.GetResultAsType("SELECT * FROM UDO2 WHERE \"Code\" = '{0}'", dc =>
                {
                    return new UDO2
                    {
                        Code = dc["Code"],
                        ColmnNum= dc["ColumnNum"],
                        ColAlias = dc["ColAlias"],
                        ColumnDesc = dc["ColumnDesc"]                        
                    };
                }, code).ToList();
                List<UDO3> udos3 = hsh.GetResultAsType("SELECT * FROM UDO3 WHERE \"Code\" = '{0}'", dc =>
                {
                    return new UDO3
                    {
                        Code = dc["Code"],
                        LineNum = dc["ColumnNum"],
                        SonNum = dc["SonNum"],
                        ColAlias = dc["ColAlias"],
                        ColDesc = dc["ColDesc"],
                        ColEdit = dataUdo(dc["ColEdit"])
                    };
                }, code).ToList();

                listaObject.Add(ObjectEnSAP);
                if (udos1.Count > 0) listaudo1.AddRange(udos1);
                if (udos2.Count > 0) listaudo2.AddRange(udos2);
                if (udos3.Count > 0) listaudo3.AddRange(udos3);
            }

            AddDataExcelUDO(listaObject, listaudo1, listaudo2, listaudo3, addon);
        }

        static void AddDataExcelUDO(List<SAPOBJECT> lista, List<UDO1> listaUdo1, List<UDO2> listaUdo2, List<UDO3> listaUdo3, string addon)
        {
            // Ruta Excel
            string path = $"D:\\SEBAS\\PlantillasAddonPE\\Excels\\{addon}\\02 Plantilla de Objetos.xlsm";

            Application excelApp = new Application();

            // Abrir el libro de trabajo
            Workbook workbook = excelApp.Workbooks.Open(path);

            // AGREGA LOS OBJETOS
            // Obtener la hoja de trabajo
            Worksheet worksheet = (Worksheet)workbook.Sheets[1]; // Hoja 1
            // Leer los datos de la hoja de trabajo
            // Recorremos la lista de datos y los escribimos en las celdas correspondientes
            for (int i = 0; i < lista.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = lista[i].Code;
                worksheet.Cells[i + 3, 2] = lista[i].Name;
                worksheet.Cells[i + 3, 3] = lista[i].TableName;
                worksheet.Cells[i + 3, 4] = lista[i].LogTable;
                worksheet.Cells[i + 3, 5] = lista[i].TYPE;
                worksheet.Cells[i + 3, 6] = lista[i].MngSeries;
                worksheet.Cells[i + 3, 7] = lista[i].CanDelete;
                worksheet.Cells[i + 3, 8] = lista[i].CanClose;
                worksheet.Cells[i + 3, 9] = lista[i].CanCancel;
                worksheet.Cells[i + 3, 10] = lista[i].ExtName;
                worksheet.Cells[i + 3, 11] = lista[i].CanFind;
                worksheet.Cells[i + 3, 12] = lista[i].CanYrTrnsf;
                worksheet.Cells[i + 3, 13] = lista[i].CanDefForm;
                worksheet.Cells[i + 3, 14] = lista[i].CanLog;
                worksheet.Cells[i + 3, 15] = lista[i].OvrWrtDll;
                worksheet.Cells[i + 3, 16] = lista[i].UIDFormat;
                worksheet.Cells[i + 3, 17] = lista[i].CanArchive;
                worksheet.Cells[i + 3, 18] = lista[i].MenuItem;
                worksheet.Cells[i + 3, 19] = lista[i].MenuCapt;
                worksheet.Cells[i + 3, 20] = lista[i].FatherMenu;
                worksheet.Cells[i + 3, 21] = lista[i].Position;
                worksheet.Cells[i + 3, 22] = lista[i].CanNewForm;
                worksheet.Cells[i + 3, 23] = lista[i].IsRebuild;
                worksheet.Cells[i + 3, 24] = lista[i].NewFormSrf;
                worksheet.Cells[i + 3, 25] = lista[i].MenuUid;
            }

            worksheet = (Worksheet)workbook.Sheets[2]; // Hoja 1

            for (int i = 0; i < listaUdo1.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = listaUdo1[i].Code;
                worksheet.Cells[i + 3, 2] = listaUdo1[i].LineNum;
                worksheet.Cells[i + 3, 3] = listaUdo1[i].TableName;
                worksheet.Cells[i + 3, 4] = listaUdo1[i].LogName;
                worksheet.Cells[i + 3, 5] = listaUdo1[i].SonName;
            }

            worksheet = (Worksheet)workbook.Sheets[3]; // Hoja 1

            for (int i = 0; i < listaUdo2.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = listaUdo2[i].Code;
                worksheet.Cells[i + 3, 2] = listaUdo2[i].ColmnNum;
                worksheet.Cells[i + 3, 3] = listaUdo2[i].ColAlias;
                worksheet.Cells[i + 3, 4] = listaUdo2[i].ColumnDesc;
            }

            worksheet = (Worksheet)workbook.Sheets[4]; // Hoja 1

            for (int i = 0; i < listaUdo3.Count; i++)
            {
                // Escribir los datos en las celdas
                worksheet.Cells[i + 3, 1] = listaUdo3[i].Code;
                worksheet.Cells[i + 3, 2] = listaUdo3[i].LineNum;
                worksheet.Cells[i + 3, 3] = listaUdo3[i].SonNum;
                worksheet.Cells[i + 3, 4] = listaUdo3[i].ColAlias;
                worksheet.Cells[i + 3, 5] = listaUdo3[i].ColDesc;
                worksheet.Cells[i + 3, 6] = listaUdo3[i].ColEdit;
            }

            // Guardar los cambios
            workbook.Save();

            // Cerrar el libro de trabajo y la aplicación Excel
            workbook.Close();
            excelApp.Quit();

        }

        static string dataUdo(string s)
        {
            return s == "Y" ? "tYES" : "tNO";
        }
    }
    public class UFD1
    { 
        public string TableID { get; set; }
        public string FieldID { get; set; }
        public string IndexID { get; set; }
        public string FldValue { get; set; }
        public string Descr { get; set; }
        public string FldDate { get; set; }
    }
    public class UDO1
    { 
        public string Code { get; set; }
        public string LineNum { get; set; }
        public string TableName { get; set; }
        public string LogName { get; set; }
        public string SonName { get; set; }
    }
    public class UDO2
    { 
        public string Code { get; set; }
        public string ColmnNum { get; set; }
        public string ColAlias { get; set; }
        public string ColumnDesc { get; set; }
    }
    public class UDO3
    { 
        public string Code { get; set; }
        public string LineNum { get; set; }
        public string SonNum { get; set; }
        public string ColAlias { get; set; }
        public string ColDesc { get; set; }
        public string ColEdit { get; set; }
    }
    public class SAPTABLE
    {
        public string tableName { get; set; }
        public string tableDescription { get; set; }
        public string tableType { get; set; }
    }
    public class SAPCOLUMN
    {
        public string Name { get; set; }
        public string FieldID { get; set; }
        public string Type { get; set; }
        public string Size { get; set; }
        public string Description { get; set; }
        public string SubType { get; set; }
        public string TableName { get; set; }
        public string EditSize { get; set; }
        public string Mandatory { get; set; }
        public string DefaultValue { get; set; }
    }
    public class SAPOBJECT
    { 
        public string Code { get; set; }
        public string Name { get; set; }
        public string TableName { get; set; }
        public string LogTable { get; set; }
        public string TYPE { get; set; }
        public string MngSeries { get; set; }
        public string CanDelete { get; set; }
        public string CanClose { get; set; }
        public string CanCancel { get; set; }
        public string ExtName { get; set; }
        public string CanFind { get; set; }
        public string CanYrTrnsf { get; set; }
        public string CanDefForm { get; set; }
        public string CanLog { get; set; }
        public string OvrWrtDll { get; set; }
        public string UIDFormat { get; set; }
        public string CanArchive { get; set; }
        public string MenuItem { get; set; }
        public string MenuCapt { get; set; }
        public string FatherMenu { get; set; }
        public string Position { get; set; }
        public string CanNewForm { get; set; }
        public string IsRebuild { get; set; }
        public string NewFormSrf { get; set; }
        public string MenuUid { get; set; }
    }
}
