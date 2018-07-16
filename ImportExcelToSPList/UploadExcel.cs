using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using MSDN.Samples.ClaimsAuth;
using System.Configuration;

namespace ImportExcelToSPList
{
    public class UploadExcel
    {
        /// <summary>
        /// Read the data from an excel file and upload it into a datatable
        /// </summary>
        public static void LoadExcelData()
        {
            string fileName = ConfigurationManager.AppSettings["ExcelFilePath"];

            //if you are using file upload control in sharepoint get the full path as follows assuming fileUpload1 is control instance
            //string fileName = fileUpload1.PostedFile.FileName

            string fileExtension = Path.GetExtension(fileName).ToUpper();
            string connectionString = "";

            if (fileExtension == ".XLS")
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fileName + "'; Extended Properties='Excel 8.0;HDR=YES;'";
            }
            else if (fileExtension == ".XLSX")
            {
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileName + "';Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            }
            if (!(string.IsNullOrEmpty(connectionString)))
            {
                string[] sheetNames = GetExcelSheetNames(connectionString);
                if ((sheetNames != null) && (sheetNames.Length > 0))
                {
                    DataTable dt = null;
                    OleDbConnection con = new OleDbConnection(connectionString);
                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheetNames[0] + "]", con);
                    dt = new DataTable();
                    da.Fill(dt);
                    InsertIntoList(dt, "ExcelList");
                }
            }
        }

        private static string[] GetExcelSheetNames(string strConnection)
        {
            var connectionString = strConnection;
            String[] excelSheets;
            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                var dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    return null;
                }
                excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
            }
            return excelSheets;
        }

        private static void InsertIntoList(DataTable listTable, string contactListName)
        {


            string siteUrl = ConfigurationManager.AppSettings["SharePointURL"]; 
            bool isBatchUpdate = Convert.ToBoolean(ConfigurationManager.AppSettings["IsBatchUpdate"]);
            
            using (Microsoft.SharePoint.Client.ClientContext context = ClaimClientContext.GetAuthenticatedContext(siteUrl))
            {
                string strListName = ConfigurationManager.AppSettings["TargetSharePointListName"];
                int counter = 0;
                Console.WriteLine(string.Format("\nContent type updation would be performed for url - {0}\n", siteUrl));
                Web site = context.Web;
                context.Load(site);
                context.Load(site.Lists, lists => lists.Include(list => list.Title, // For each list, retrieve Title and Id. 
                        list => list.Id));
                context.ExecuteQuery();
                Microsoft.SharePoint.Client.List GlossaryList = site.Lists.GetByTitle(strListName); 
                for (int iRow = 0; iRow < listTable.Rows.Count; iRow++)
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem oListItem = GlossaryList.AddItem(itemCreateInfo);
                    //the column name below is the absolute name of the column which can be seen in List Settings-->List columns
                    //Modify the below rows as per your target list columns
                    oListItem["Title"] = Convert.ToString(listTable.Rows[iRow][0]); //List Title
                    oListItem["r9dw"] = Convert.ToString(listTable.Rows[iRow][2]); //Definition
                    oListItem["gpqr"] = Convert.ToString(listTable.Rows[iRow][3]); //TermCategory
                    counter++;
                    oListItem.Update();
                    if (isBatchUpdate == true)
                    {                        
                        if ((counter % 10) == 0) //for batch batch by batch updation in multiples of 10
                        {
                            context.ExecuteQuery();
                            counter = 0; //reset the counter
                        }
                    }
                    else
                        context.ExecuteQuery(); //Either use this if you want to update every item one by one
                    
                }
            }

        }

    }
}