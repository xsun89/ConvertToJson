using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using Spire.Xls;
using Formatting = Newtonsoft.Json.Formatting;
using Workbook = Spire.Xls.Workbook;
using Worksheet = Spire.Xls.Worksheet;


namespace CovertToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Workbook workbook = new Workbook();

                workbook.LoadFromFile(@"C:\temp\CFRI_Clinical_DB_Approved_Studies_v3.xlsx");
                //Initailize worksheet
                Worksheet sheet = workbook.Worksheets[0];

                DataTable dataTable = sheet.ExportDataTable();

                DataTableToJsonObj(dataTable);


                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadKey();
            }
        }
        public static void DataTableToJsonObj(DataTable dt)
        {
            using (StreamWriter outputFile = new StreamWriter(@"C:\temp\jsonFinal_v2.txt", false))
            {


                DataSet ds = new DataSet();
                ds.Merge(dt);
                StringBuilder JsonString = new StringBuilder();
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    var rowCount = ds.Tables[0].Rows.Count;
                    JsonString.Append("[");
                    XmlDocument doc = new XmlDocument();
                    for (int i = 0; i < rowCount; i++)
                    {
                        JsonString.Append("{");
                        for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                        {
                            if (
                                ds.Tables[0].Columns[j].ColumnName.ToString() == "Types of Funds"
                                || ds.Tables[0].Columns[j].ColumnName.ToString() == "Institutions and Sites"
                                ||
                                ds.Tables[0].Columns[j].ColumnName.ToString() == "Advertisement to Recruit Participants"
                                || ds.Tables[0].Columns[j].ColumnName.ToString() == "Consent Forms"
                                || ds.Tables[0].Columns[j].ColumnName.ToString() == "Assent Forms"
                                || ds.Tables[0].Columns[j].ColumnName.ToString() == "COI"
                                )
                            {
                                var xmlData = ds.Tables[0].Rows[i][j].ToString();
                                string jsonData = null;
                                if (!String.IsNullOrEmpty(xmlData))
                                {
                                    doc.LoadXml(xmlData);
                                    jsonData = JsonConvert.SerializeXmlNode(doc, Formatting.None, true);
                                    JsonString.Append(jsonData.Remove(0, 1) + ",")
                                        .Remove(JsonString.ToString().LastIndexOf('}'), 1);


                                }
                            }
                            else if (ds.Tables[0].Columns[j].ColumnName.ToString() == "PI"
                                     || ds.Tables[0].Columns[j].ColumnName.ToString() == "Primary Contact"
                                )
                            {
                                var xmlData = ds.Tables[0].Rows[i][j].ToString();
                                string jsonText = null;
                                if (!String.IsNullOrEmpty(xmlData))
                                {
                                    doc.LoadXml(xmlData);
                                    jsonText = JsonConvert.SerializeXmlNode(doc, Formatting.None, true);
                                    JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString().Trim() +
                                                      "\":" +
                                                      jsonText.Trim() + ",").Replace("@userID", "userID")
                                        .Replace("@name", "name")
                                        .Replace("@email", "email")
                                        .Replace("@phone", "phone")
                                        .Replace("@site", "site")
                                        .Replace("@rank", "rank")
                                        .Replace("@location", "location");



                                }
                            }
                            else
                            {
                                var data =
                                    ds.Tables[0].Rows[i][j].ToString().Replace("\"", "\\\""); 
                                        
                                if (j < ds.Tables[0].Columns.Count - 1)
                                {
                                    JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString().Trim() +
                                                      "\":" + "\"" +
                                                      data + "\",");
                                }
                                else if (j == ds.Tables[0].Columns.Count - 1)
                                {

                                    JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString().Trim() +
                                                      "\":" + "\"" +
                                                      data + "\"");
                                }
                            }
                        }
                        if (i == ds.Tables[0].Rows.Count - 1)
                        {
                            JsonString.Append("}");
                        }
                        else
                        {
                            JsonString.Append("},");
                        }
                        outputFile.WriteLine(JsonString.ToString());

                        JsonString.Clear();

                    }
                    JsonString.Append("]");

                    outputFile.WriteLine(JsonString.ToString());
                }
            }
        }
    }
}
