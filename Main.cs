using System;
using System.Data;
using System.Windows.Forms;
using SDKWrapper;
using System.Configuration;
using System.IO;
using Excel;
using System.Collections.Generic;
using SimpleLogger;

namespace eArchive
{
    public partial class Main : Form
    {
        FarzinMethods SDKmethod = new FarzinMethods();
        string ExcelFilePath = ConfigurationManager.AppSettings["ExcelFilePath"];
        string FarzinUserName = ConfigurationManager.AppSettings["FarzinUserName"];
        string FarzinHashPassword = ConfigurationManager.AppSettings["FarzinHashPassword"];
        string LogDirectoryFolder = ConfigurationManager.AppSettings["LogDirectoryFolder"];

        public Main()
        {
            InitializeComponent();
            SimpleLog.SetLogFile(logDir: LogDirectoryFolder, writeText: true, check: false);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable ExcelCleanData = FillDatatableFromExcel(ExcelFilePath, 0);
                ImportFormsFromDatatable(ExcelCleanData);
                foreach (string info in SDKmethod.errorList)
                {
                    SimpleLog.Info(info);
                }
            }
            catch (Exception ex)
            {
                SimpleLog.Info(ex.Message);
                return;
            }
        }

        private void ImportFormsFromDatatable(DataTable FormsDataSource)
        {
            List<string> DocsStructure = new List<string>();

            //create a datatable for form's fiels
            DataTable FieldsTable = new DataTable();
            FieldsTable.Columns.AddRange(new DataColumn[] {
                                        new DataColumn("FieldName", typeof(String)),
                                        new DataColumn("FieldType", typeof(String)),
                                        new DataColumn("FieldValue", typeof(String)),
                                     });

            foreach (DataRow row in FormsDataSource.Rows)
            {
                //Set form's fiels
                //row[0] = skip this because of first column of excel's file
                FieldsTable.Rows.Add("ID", "int", row[1].ToString());
                FieldsTable.Rows.Add("code", "nvarchar", row[2].ToString());
                FieldsTable.Rows.Add("name", "nvarchar", row[3].ToString());
                FieldsTable.Rows.Add("scaned", "nvarchar", row[4].ToString());
                FieldsTable.Rows.Add("registerDate", "nvarchar", row[5].ToString());
                FieldsTable.Rows.Add("requestYear", "nvarchar", row[6].ToString());
                FieldsTable.Rows.Add("clause", "nvarchar", row[7].ToString());
                FieldsTable.Rows.Add("description", "nvarchar", row[8].ToString());
                //create a list of XML of form's structure
                DocsStructure.Add(SDKmethod.BuildDocStructure("AgricultureArchive", "Entity_AgricultureArchive", "211", "333", FieldsTable));
                FieldsTable.Clear();
            }

            //Insert forms depend on DocsStructure list
            SDKmethod.Insert(DocsStructure, FarzinUserName, FarzinHashPassword);
        }

        private DataTable FillDatatableFromExcel(string ExcelPath, int PrimaryKeyColumnIndex)
        {
            DataSet ExcelDataSet;
            DataRow[] ExcelFilteredRows;
            DataTable ExcelFilteredDataTable;
            FileStream stream = File.Open(ExcelPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;
            ExcelDataSet = excelReader.AsDataSet();
            ExcelFilteredRows = ExcelDataSet.Tables[0].Select("[" + ExcelDataSet.Tables[0].Columns[PrimaryKeyColumnIndex].ColumnName + "] IS NOT NULL");
            ExcelFilteredDataTable = ExcelFilteredRows.CopyToDataTable();
            ExcelDataSet.Clear();
            ExcelDataSet.Dispose();
            excelReader.Close();
            return ExcelFilteredDataTable;
        }


        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        private void button2_Click(object sender, EventArgs e)
        {
            string sourcePath = @"D:\scans\";
            DirSearch(sourcePath);
        }

        private void DirSearch(string sDir)
        {
            string CleanFolderName = null;

            try
            {
                foreach (string file in Directory.GetFiles(sDir))
                {
                    lstFilesFound.Items.Add(file);
                }
                
                foreach (string folder in Directory.GetDirectories(sDir))
                {
                    CleanFolderName = folder.Replace(sDir, null);
                    FindClassificationCode(CleanFolderName);
                    DirSearch(folder);
                }
            }
            catch (Exception ex)
            {
                listBox1.Items.Add(ex.Message);
            }
        }



        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private string FindClassificationCode(string FolderName)
        {
            return FolderName.Remove(0, 11);
        }
    }
}
