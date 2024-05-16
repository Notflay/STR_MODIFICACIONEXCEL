using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Office.Interop.Excel;

namespace ModificacionExcvel
{
    public class Program
    {
        static void Main()
        {
            // Obtiene Tabla
            ObtieneDataTable();

            // Obtiene Columnas
            ObtieneDataColumns();

        }

        static void ObtieneDataColumns()
        {
            XmlDocument xmlDoc = new XmlDocument();
            // C:\Users\Sebastian\source\repos\STR_ADDONPERU_INSTALADOR\STR_ADDONPERU_INSTALADOR\Resources\Localizacion\UT.vte
            string path = "D:\\Chamba Backend\\AddonPeru\\PlantillasImplementacion\\XML\\Localizacion\\UF.xml";
            // Cargar el archivo XML con la codificación UTF-16
            using (XmlReader reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
            {
                xmlDoc.Load(reader);
            }

            XmlNodeList tableNodes = xmlDoc.SelectNodes("//BO/UserFieldsMD/row");

            List<SAPCOLUMN> listaColumns = new List<SAPCOLUMN>();

            HanaADOHelper hsh = new HanaADOHelper();
            foreach (XmlNode tableNode in tableNodes)
            {
                SAPCOLUMN col = new SAPCOLUMN();
                // Obtener los valores de TableName, TableDescription y TableType
                col.Name = tableNode.SelectSingleNode("Name").InnerText;
                // Obtiene Data de Columnas
                col.FieldID = hsh.insertValueSql("SELECT \"FieldID\" FROM CUFD WHERE \"AliasID\" = '{0}' and \"TableID\" = '{1}'", tableNode.SelectSingleNode("Name").InnerText, tableNode.SelectSingleNode("TableName").InnerText);
                col.Type = tableNode.SelectSingleNode("Type").InnerText;
                col.Size = tableNode.SelectSingleNode("Size").InnerText;
                col.Description = tableNode.SelectSingleNode("Description").InnerText;
                col.SubType = tableNode.SelectSingleNode("SubType").InnerText;
                col.TableName = tableNode.SelectSingleNode("TableName").InnerText;
                col.EditSize = tableNode.SelectSingleNode("EditSize").InnerText;
                col.Mandatory = tableNode.SelectSingleNode("Mandatory").InnerText;
                col.DefaultValue = tableNode.SelectSingleNode("DefaultValue") == null ? "" : tableNode.SelectSingleNode("DefaultValue").InnerText;
                listaColumns.Add(col);
            }
            AddDataExcelCol(listaColumns);
        }

        static void AddDataExcelCol(List<SAPCOLUMN> lista)
        {
            // Ruta Excel
            string path = "D:\\Chamba Backend\\AddonPeru\\PlantillasImplementacion\\01 Plantilla de Campos.xlsm";

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
                worksheet.Cells[i + 3, 2] = lista[i].TableName;
                worksheet.Cells[i + 3, 3] = lista[i].Name;
                worksheet.Cells[i + 3, 4] = lista[i].Type;
                worksheet.Cells[i + 3, 5] = lista[i].Size;
                worksheet.Cells[i + 3, 6] = lista[i].Description;
                worksheet.Cells[i + 3, 7] = lista[i].SubType;
                worksheet.Cells[i + 3, 9] = lista[i].DefaultValue;
                worksheet.Cells[i + 3, 10] = lista[i].EditSize;
                worksheet.Cells[i + 3, 11] = lista[i].Mandatory;
            }

            // Guardar los cambios
            workbook.Save();

            // Cerrar el libro de trabajo y la aplicación Excel
            workbook.Close();
            excelApp.Quit();

        }

        static void ObtieneDataTable()
        {
            XmlDocument xmlDoc = new XmlDocument();
            // C:\Users\Sebastian\source\repos\STR_ADDONPERU_INSTALADOR\STR_ADDONPERU_INSTALADOR\Resources\Localizacion\UT.vte
            string path = "D:\\Chamba Backend\\AddonPeru\\PlantillasImplementacion\\XML\\Localizacion\\UT.xml";
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

            AddDataExcelTable(listTables);
        }

        static void AddDataExcelTable(List<SAPTABLE> lista)
        {
            // Ruta Excel
            string path = "D:\\Chamba Backend\\AddonPeru\\PlantillasImplementacion\\00 Plantilla de Tablas.xlsm";

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
}
