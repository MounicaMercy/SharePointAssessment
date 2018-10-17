using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.SharePoint.Client;
using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using File = Microsoft.SharePoint.Client.File;
using System.Data;
using DocumentFormat.OpenXml;

using System.Runtime.InteropServices;

namespace SharePointAssessment
{
    class FileInformations
    {
        static System.Data.DataTable dt = new System.Data.DataTable();
        static void Main(string[] args)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            //Console.WriteLine("Enter User Name");
            string UserName = "mounika.pasupunuri@acuvate.com";
            Console.WriteLine("Enter Your Password");
            SecureString Password = GetPassword();
            using (ClientContext clientcontext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/Assessment"))
            {
                //Authentication();
                clientcontext.Credentials = new SharePointOnlineCredentials(UserName, Password);

                string FileToRead = "FileInformation.xlsx";
                try
                {
                    ReadExcelSheet(clientcontext, FileToRead);
                    
                    isError = false;
                }
                catch (Exception e)
                {
                    isError = true;
                    strErrorMsg = e.Message;
                }
                finally
                {
                    if (isError)
                    {
                        //Logging
                    }
                }
            }
        }
        private static SecureString GetPassword()
            {
                ConsoleKeyInfo info;
                //Get the user's password as a SecureString  
                SecureString securePassword = new SecureString();
                do
                {
                    info = Console.ReadKey(true);
                    if (info.Key != ConsoleKey.Enter)
                    {
                        securePassword.AppendChar(info.KeyChar);
                    }
                }
                while (info.Key != ConsoleKey.Enter);
                return securePassword;
            }
        private static void ReadExcelSheet(ClientContext clientcontext,string filename)
        {
            Web web = clientcontext.Web;
            List list = web.Lists.GetByTitle("AssessmentDoc");
            clientcontext.Load(list, lists => lists.RootFolder);
            clientcontext.ExecuteQuery();

            String FileName = (list.RootFolder.ServerRelativeUrl + "/"+ filename);
            File file = clientcontext.Web.GetFileByServerRelativeUrl(FileName);
            ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
            clientcontext.Load(file);
            clientcontext.ExecuteQuery();
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientcontext, document, cell);
                                dt.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dt.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientcontext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dt.Rows.Add(dataRow);
                                }
                            }
                            dt.Rows.RemoveAt(0);
                        }
                    }
                    //UpdateExcelSheet(clientcontext,dt, filename);
                }
                UploadFile(clientcontext, dt);             //Uploading File
                foreach (DataRow dataRow in dt.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        Console.WriteLine(item);
                    }
                }
                //Console.WriteLine(dt.Rows[0].Field<string>(4));
                Console.ReadKey();
            }
        }
        private static string GetCellValue(ClientContext clientcontext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
            return value;
        }
        private static void UpdateExcelSheet(ClientContext clientcontext,DataTable dataTable,string filename)
        {

        }
        //private static void UpdateSPList(ClientContext clientcontext, DataTable dataTable, string fileName)
        //{
        //    bool isError = true;
        //    string strErrorMsg = string.Empty;
        //    Int32 count = 0;
        //    const string lstName = "EmployeesData";
        //    const string lstColTitle = "Title";
        //    const string lstColAddress = "Address";
        //    try
        //    {
        //        string fileExtension = ".xlsx";
        //        string fileNameWithOutExtension = fileName.Substring(0, fileName.Length - fileExtension.Length);
        //        if (fileNameWithOutExtension.Trim() == lstName)
        //        {
        //            SP.List oList = clientcontext.Web.Lists.GetByTitle(fileNameWithOutExtension);
        //            foreach (DataRow row in dataTable.Rows)
        //            {
        //                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
        //                ListItem oListItem = oList.AddItem(itemCreateInfo);
        //                oListItem[lstColTitle] = row[0];
        //                oListItem[lstColAddress] = row[1];
        //                oListItem.Update();
        //                clientcontext.ExecuteQuery();
        //                count++;
        //            }
        //        }
        //        else
        //        {
        //            count = 0;
        //        }
        //        if (count == 0)
        //        {
        //            Console.Write("Error: List: '" + fileNameWithOutExtension + "' is not found in SharePoint.");
        //        }
        //        isError = false;
        //    }
        //    catch (Exception e)
        //    {
        //        isError = true;
        //        strErrorMsg = e.Message;
        //    }
        //    finally
        //    {
        //        if (isError)
        //        {
        //            //Logging
        //        }
        //    }
        //}
        public static void UploadFile(ClientContext clientcontext, DataTable dt)
        {
            bool flag = true;
            long bytes = 0;
            String Reason = "";
            String UploadStatus = "";

            Web web = clientcontext.Web;
            
            List docs = clientcontext.Web.Lists.GetByTitle("Files");
            clientcontext.ExecuteQuery();

            foreach (DataRow row in dt.Rows)
            {
                string path = row.Field<string>("File Path");    //Getting the coulmn-file path from data table
                string fileName = Path.GetFileName(path);          //Getting the file name from the path

                string CreatedBy = row[3].ToString();

                System.IO.FileInfo fileInfo = new System.IO.FileInfo(path);
                string FileType = fileInfo.Extension;

                try
                {
                   
                    if (fileInfo.Exists)                                      //Checking if file exists or not
                        bytes = fileInfo.Length;                                  //read the file size
                    else                                                         // if the file not exit then status will be failed
                    {
                        UploadStatus = "Failed";
                        Reason = "File does not exists on given path";
                        flag = false;
                    }
                    if (flag == true && (bytes < 10000000))                      //checking the file size before upload
                    {
                       
                        //int depart = Convert.ToInt32(row[2]);

                        FileCreationInformation file = new FileCreationInformation();
                        file.Content = System.IO.File.ReadAllBytes(path);
                        file.Url = fileName;
                        file.Overwrite = true;
                        File fileToUpload = docs.RootFolder.Files.Add(file);
                        clientcontext.Load(fileToUpload);
                        clientcontext.ExecuteQuery();
                        Console.WriteLine("File Uploaded!!");

                        ListItem ItemToBeUpdated = docs.GetItemById(1);
                        ItemToBeUpdated["Created By"] = CreatedBy.ToString();
                        ItemToBeUpdated["Type Of the File"] = FileType.ToString();
                        ItemToBeUpdated.Update();
                        clientcontext.ExecuteQuery();
                    }
                    else if (flag)
                    {
                        // Console.WriteLine("File Size is High");
                       
                        UploadStatus = "Failed";
                        Reason = "File size is to high";
                    }
                }

                catch (Exception e)
                {
                    Console.WriteLine("File Size is High");
                    //Console.WriteLine(e.Message);
                }
            }
        }
    }
}



