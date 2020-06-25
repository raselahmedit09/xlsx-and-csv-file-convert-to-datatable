using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var connectionString = ConfigurationManager.ConnectionStrings["JulyDB"].ConnectionString;

            //create instanace of database connection
            SqlConnection conn = new SqlConnection(connectionString);
            //open connection
            conn.Open();

            phoneCallActivity(conn);
            EmailActivity(conn);

            //Console.Read();
            conn.Close(); // connection close 
        }

        private static void phoneCallActivity(SqlConnection conn)
        {
            string[] filePaths = Directory.GetFiles(@"D:\Rasel Ahmed\UN Files\Phone Call\", "*.*");
            foreach (string filePath in filePaths)
            {
                Microsoft.Office.Interop.Excel.Application objXL = null;
                Microsoft.Office.Interop.Excel.Workbook objWB = null;
                objXL = new Microsoft.Office.Interop.Excel.Application();
                objWB = objXL.Workbooks.Open(filePath);

                DataTable dt = new DataTable();

                for (int ws = 1; ws <= objWB.Sheets.Count; ws++)
                {
                    Microsoft.Office.Interop.Excel.Worksheet objSHT = objWB.Worksheets[ws];

                    int rows = (ws == objWB.Sheets.Count) ? objSHT.UsedRange.Rows.Count - 1 : objSHT.UsedRange.Rows.Count;
                    int cols = objSHT.UsedRange.Columns.Count;

                    //DataTable dt = new DataTable();

                    int columnHeader = 0;
                    for (int i = 1; i <= cols; i++)
                    {
                        string rowVlaue = objSHT.Cells[i, 1].Text;

                        if (!string.IsNullOrEmpty(rowVlaue))
                        {
                            columnHeader = i;
                            break;
                        }
                    }

                    if (ws == 1) // get column name from 1st sheet.
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            string colname = objSHT.Cells[columnHeader, c].Text;
                            if (colname.Length > 0) // check empty columns
                            {
                                dt.Columns.Add(colname.Replace(" ", "").ToLower());
                            }
                        }
                    }

                    for (int r = columnHeader + 1; r <= rows; r++) // row data 
                    {
                        DataRow dr = dt.NewRow();
                        for (int c = 1; c <= cols; c++)
                        {
                            string colname = objSHT.Cells[columnHeader, c].Text;
                            if (colname.Length > 0) // check empty columns
                            {
                                dr[c - 1] = objSHT.Cells[r, c].Text;
                            }
                        }

                        // check emplty row
                        string firstColumnValue = dr.ItemArray[0].ToString();
                        string secondColumnValue = dr.ItemArray[1].ToString();

                        if (firstColumnValue.Length > 0 || secondColumnValue.Length > 0) // check empty row by fist two value. 
                        {
                            dt.Rows.Add(dr);
                        }
                    }

                    if (ws == objWB.Sheets.Count)
                    {
                        System.Data.DataColumn deviceName = new System.Data.DataColumn("devicename", typeof(System.String));
                        deviceName.DefaultValue = objSHT.Cells[objSHT.UsedRange.Rows.Count, 1].Text;
                        dt.Columns.Add(deviceName);

                        System.Data.DataColumn dateCreated = new System.Data.DataColumn("activitydate", typeof(System.String));
                        var dateText = objSHT.Cells[objSHT.UsedRange.Rows.Count, 2].Text;
                        string activityDate = dateText.Substring(0, dateText.IndexOf("-")).Trim();
                        dateCreated.DefaultValue = DateTime.Parse(activityDate);
                        dt.Columns.Add(dateCreated);
                    }
                }

                // insert data to database 
                foreach (DataRow row in dt.Rows)
                {
                    string queryString = @"INSERT INTO PhoneSystemImportActivityData 
                                        (ContactId,	
                                        AgentId,	
                                        AgentName,	
                                        ShiftDuration,	
                                        ACDCount,	
                                        ACDTalkTime,	
                                        NonACDCount,	
                                        NonACDTalkTime,	
                                        OutboundCount,	
                                        OutboundTalkTime,	
                                        MakeBusyTime,	
                                        DeviceName,	
                                        ActivityDate)
                                        VALUES([dbo].[fnGetContactIdByFullName]('" + row["agentname"].ToString() + "')," + Convert.ToInt32(row["agentid"]) + ", '" + row["agentname"].ToString() + "' " +
                                        ",'" + row["shiftduration"].ToString() + "' ," + Convert.ToInt32(row["acdcount"]) + "" +
                                        ",'" + row["acdtalktime"].ToString() + "'," + Convert.ToInt32(row["nonacdcount"]) + "" +
                                        ",'" + row["nonacdtalktime"].ToString() + "'," + Convert.ToInt32(row["outboundcount"]) + "" +
                                        ",'" + row["outboundtalktime"].ToString() + "','" + row["makebusytime"].ToString() + "'" +
                                        ",'" + row["devicename"].ToString() + "','" + row["activitydate"].ToString() + "'" +
                                        ")";


                    SqlCommand command = new SqlCommand(queryString, conn);
                    command.ExecuteNonQuery();
                }

                objWB.Close();
                objXL.Quit();

                // move file to archive folder
                moveFile(filePath);
            }
        }

        private static void EmailActivity(SqlConnection conn)
        {
            string[] filePaths = Directory.GetFiles(@"D:\Rasel Ahmed\UN Files\Email\", "*.*");
            foreach (string filePath in filePaths)
            {
                //reading all the lines(rows) from the file.
                string[] rows = File.ReadAllLines(filePath);

                DataTable dtData = new DataTable();
                string[] rowValues = null;
                DataRow dr = dtData.NewRow();

                //Creating columns
                if (rows.Length > 0)
                {
                    foreach (string columnName in rows[0].Split(','))
                        dtData.Columns.Add(columnName.Replace(" ", "").ToLower());
                }

                //Creating row for each line.(except the first line, which contain column names)
                for (int row = 1; row < rows.Length; row++)
                {
                    rowValues = rows[row].Split(',');
                    dr = dtData.NewRow();
                    dr.ItemArray = rowValues;
                    dtData.Rows.Add(dr);
                }

                // save data into database table 
                foreach (DataRow row in dtData.Rows)
                {
                    int isDeleted = row["isdeleted"].ToString().ToLower() == "false" ? 0 : 1;
                    string queryString = @"INSERT INTO EmailSystemImportActivityData 
                                            (ContactId
                                            ,ReportRefreshDate
                                            ,UserPrincipalName
                                            ,DisplayName
                                            ,IsDeleted
                                            ,DeletedDate
                                            ,LastActivityDate
                                            ,SendCount
                                            ,ReceiveCount
                                            ,ReadCount
                                            ,MeetingCreatedCount
                                            ,MeetingInteractedCount
                                            ,AssignedProducts
                                            ,ReportPeriod
                                            ,InsertedDate
                                        )
                                        
                                        VALUES([dbo].[fnGetContactIdByFullName]('" + row["displayname"].ToString() + "')" +
                                            ",'" + row["reportrefreshdate"].ToString() + "'" +
                                            ",'" + row["userprincipalname"].ToString() + "' " +
                                            ",'" + row["displayname"].ToString() + "'" +
                                            "," + isDeleted + "" +
                                            ",'" + row["deleteddate"].ToString() + "'" +
                                            ",'" + row["lastactivitydate"].ToString() + "'" +
                                            "," + Convert.ToInt32(row["sendcount"]) + "" +
                                            "," + Convert.ToInt32(row["receivecount"]) + "" +
                                            "," + Convert.ToInt32(row["readcount"]) + "" +
                                            "," + Convert.ToInt32(row["meetingcreatedcount"]) + "" +
                                            "," + Convert.ToInt32(row["meetinginteractedcount"]) + "" +
                                            ",'" + row["assignedproducts"].ToString() + "'" +
                                            "," + Convert.ToInt32(row["reportperiod"]) + "" +
                                            ",getDate()" +
                                        ")";


                    SqlCommand command = new SqlCommand(queryString, conn);
                    command.ExecuteNonQuery();
                }

                // move file to archive folder
                moveFile(filePath);

            }
        }

        private static void moveFile(string filePath)
        {
            string sourceDirecotryName = Path.GetDirectoryName(filePath);
            string destinationDirectoryName = Path.Combine(sourceDirecotryName, "Archive", DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString());
            if (!Directory.Exists(destinationDirectoryName))
            {
                Directory.CreateDirectory(destinationDirectoryName);
            }

            string destinationPath = Path.Combine(destinationDirectoryName, DateTime.Now.Day.ToString() + "_" + Path.GetFileName(filePath));
            File.Move(filePath, destinationPath);
        }
    }
}
