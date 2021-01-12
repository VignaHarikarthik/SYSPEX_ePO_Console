using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SYSPEX_ePO_Console
{
    class Program
    {
        #region ***** SQL Connection*****
        static readonly SqlConnection SGConnection = new SqlConnection("Server=192.168.1.21;Database=SYSPEX_LIVE;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection JBConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Technologies (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection JKConnection = new SqlConnection("Server=192.168.1.21;Database=PT SYSPEX KEMASINDO;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection SBConnection = new SqlConnection("Server=192.168.1.21;Database=PT SYSPEX MULTITECH;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection KLConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Mechatronic (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection PGConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Industries (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static SqlConnection SAPCon12 = new SqlConnection("Server=192.168.1.21;Database=AndriodAppDB;Uid=Sa;Pwd=Password1111;");
        static string SQLQuery;
        #endregion


        static void Main(string[] args)
        {
            EPO("65ST"); // Go Live on 31/08/20
            System.Threading.Thread.Sleep(5000);
            EPO("04SI"); // Go Live on 17/06/20
            System.Threading.Thread.Sleep(5000);
            EPO("03SM"); // Go Live on 17/06/20
            System.Threading.Thread.Sleep(5000);
            EPO("07ST"); //Go Live on 22/06/20


        }

        private static void EPO(string CompanyCode)
        {
            bool flag;

            DataSet ds = GetPoNos(CompanyCode);
            if (ds.Tables[0].Rows.Count > 0)

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {

                    flag = SendInvociePDF(ds.Tables[0].Rows[i]["docnum"].ToString(),
                                          ds.Tables[0].Rows[i]["docentry"].ToString(),
                                          ds.Tables[0].Rows[i]["E_Mail"].ToString(),
                                          ds.Tables[0].Rows[i]["cc"].ToString(), CompanyCode, ds.Tables[0].Rows[i]["cardname"].ToString());
                    DataTable dt = CheckDuplicateLog(ds.Tables[0].Rows[i]["docnum"].ToString(), CompanyCode).Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        if (flag == true)
                        {
                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "UPDATE syspex_ePO Set Status = '1'  where DocNum ='" + ds.Tables[0].Rows[i]["docnum"].ToString() + "' and Company ='" + CompanyCode + "'";
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();
                        }
                        else
                        {
                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "UPDATE syspex_ePO Set Status = '0'  where DocNum ='" + ds.Tables[0].Rows[i]["docnum"].ToString() + "' and Company ='" + CompanyCode + "'";
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();
                        }

                    }
                    else
                    {

                        if (flag == true)
                        {

                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = @"INSERT INTO syspex_ePO(Company,DocNum, CustomerCode,CustomerName,ToEmail,Path,SendDate,DocDate,Time,Status,CC) 
                            VALUES(@param1,@param2,@param4,@param5 ,@param6,@param7,@param8,@param9,@param10,@param11,@param12)";
                            cmd.Parameters.AddWithValue("@param1", CompanyCode);
                            cmd.Parameters.AddWithValue("@param2", ds.Tables[0].Rows[i]["docnum"].ToString());
                            cmd.Parameters.AddWithValue("@param4", ds.Tables[0].Rows[i]["CardCode"].ToString());
                            cmd.Parameters.AddWithValue("@param5", ds.Tables[0].Rows[i]["cardname"].ToString());
                            cmd.Parameters.AddWithValue("@param6", ds.Tables[0].Rows[i]["E_Mail"].ToString());
                            cmd.Parameters.AddWithValue("@param7", "F:\\ePo\\" + CompanyCode + "\\" + ds.Tables[0].Rows[i]["docnum"].ToString() + ".pdf");
                            cmd.Parameters.AddWithValue("@param8", DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
                            cmd.Parameters.AddWithValue("@param9", ds.Tables[0].Rows[i]["docdate"].ToString());
                            cmd.Parameters.AddWithValue("@param10", DateTime.Parse(DateTime.Now.TimeOfDay.ToString()));
                            cmd.Parameters.AddWithValue("@param11", "1");
                            cmd.Parameters.AddWithValue("@param12", ds.Tables[0].Rows[i]["CC"].ToString());
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();

                        }
                        else
                        {
                            SqlCommand cmd = new SqlCommand();
                            SAPCon12.Open();
                            cmd.Connection = SAPCon12;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = @"INSERT INTO syspex_ePO(Company,DocNum, CustomerCode,CustomerName,ToEmail,Path,SendDate,DocDate,Time,Status,CC) 
                            VALUES(@param1,@param2,@param4,@param5 ,@param6,@param7,@param8,@param9,@param10,@param11,@param12)";
                            cmd.Parameters.AddWithValue("@param1", CompanyCode);
                            cmd.Parameters.AddWithValue("@param2", ds.Tables[0].Rows[i]["docnum"].ToString());
                            cmd.Parameters.AddWithValue("@param4", ds.Tables[0].Rows[i]["CardCode"].ToString());
                            cmd.Parameters.AddWithValue("@param5", ds.Tables[0].Rows[i]["cardname"].ToString());
                            cmd.Parameters.AddWithValue("@param6", ds.Tables[0].Rows[i]["E_Mail"].ToString());
                            cmd.Parameters.AddWithValue("@param7", "F:\\ePO\\" + CompanyCode + "\\" + ds.Tables[0].Rows[i]["docnum"].ToString() + ".pdf");
                            cmd.Parameters.AddWithValue("@param8", DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
                            cmd.Parameters.AddWithValue("@param9", ds.Tables[0].Rows[i]["docdate"].ToString());
                            cmd.Parameters.AddWithValue("@param10", DateTime.Parse(DateTime.Now.TimeOfDay.ToString()));
                            cmd.Parameters.AddWithValue("@param11", "0");
                            cmd.Parameters.AddWithValue("@param12", ds.Tables[0].Rows[i]["CC"].ToString());
                            cmd.ExecuteNonQuery();
                            SAPCon12.Close();
                        }
                    }
                }
        }

        private static bool SendInvociePDF(string DocNum, string DocEntry, string To, string CC, string CompanyCode, string VendorName)
        {
            bool success;
            string Databasename = "";
            //To = "vigna@syspex.com";

            if (CompanyCode == "65ST")
                Databasename = "SYSPEX_LIVE";
            if (CompanyCode == "03SM")
                Databasename = "Syspex Mechatronic (M) Sdn Bhd";
            if (CompanyCode == "07ST")
                Databasename = "Syspex Technologies (M) Sdn Bhd";
            if (CompanyCode == "21SK")
                Databasename = "PT SYSPEX KEMASINDO";
            if (CompanyCode == "31SM")
                Databasename = "PT SYSPEX MULTITECH";
            if (CompanyCode == "04SI")
                Databasename = "Syspex Industries (M) Sdn Bhd";

            try
            {

                ReportDocument cryRpt = new ReportDocument();

                if ((CompanyCode == "03SM") || (CompanyCode == "04SI"))
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_03SM&04SI.rpt");

                if ((CompanyCode == "21SK") || (CompanyCode == "31SM"))
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_21SK&31SM.rpt");

                if (CompanyCode == "07ST")
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_07ST.rpt");

                if (CompanyCode == "65ST")
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_65ST.rpt");

                new TableLogOnInfos();
                TableLogOnInfo crtableLogoninfo;
                var crConnectionInfo = new ConnectionInfo();

                ParameterFieldDefinitions crParameterFieldDefinitions;
                ParameterFieldDefinition crParameterFieldDefinition;
                ParameterValues crParameterValues = new ParameterValues();
                ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

                crParameterDiscreteValue.Value = Convert.ToString(DocEntry);
                crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
                crParameterFieldDefinition = crParameterFieldDefinitions["@DOCENTRY"];
                crParameterValues = crParameterFieldDefinition.CurrentValues;

                crParameterValues.Clear();
                crParameterValues.Add(crParameterDiscreteValue);
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

                crConnectionInfo.ServerName = "SYSPEXSAP04";
                crConnectionInfo.DatabaseName = Databasename;
                crConnectionInfo.UserID = "sa";
                crConnectionInfo.Password = "Password1111";

                var crTables = cryRpt.Database.Tables;
                foreach (Table crTable in crTables)
                {
                    crtableLogoninfo = crTable.LogOnInfo;
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                    crTable.ApplyLogOnInfo(crtableLogoninfo);
                }



                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();

                CrDiskFileDestinationOptions.DiskFileName = "F:\\ePO\\" + CompanyCode + "\\" + DocNum + ".pdf";
                CrExportOptions = cryRpt.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                cryRpt.Export();

                //// Email Part 

                MailMessage mm = new MailMessage
                {
                    From = new MailAddress("noreply@syspex.com")
                };


                foreach (var address in To.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (IsValidEmail(address) == true)
                    {
                        mm.To.Add(address);
                    }
                }

                mm.IsBodyHtml = true;
                mm.Subject = VendorName + " " + ": " + "PO#" + DocNum;
                if (CompanyCode != "65ST")
                {
                    //mm.Subject = " Purchase Order No:" + DocNum;
                    mm.Body = "<p>Dear Valued Supplier,</p> <p>Attached please find our PO number " + DocNum + ", if you have any questions please call us immediately.</p>" +
                        "<p> Regards,</p>" +
          "<p> Procurement Team</p> ";


                }
                else
                {
                    mm.Body = ST_HTMLBULIDER(DocNum);
                    //  mm.To.Add("vigna@syspex.com");
                }


                foreach (var address in CC.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).Distinct())
                {
                    mm.CC.Add(new MailAddress(address)); //Adding Multiple CC email Id
                }


                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    EnableSsl = true
                };
                System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential
                {
                    UserName = "noreply@syspex.com",
                    Password = "design35"
                };
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = 587;
                mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                smtp.Send(mm);
                success = true;


            }
            catch (CrystalReportsException ex)
            {

                throw ex;
            }

            return success;
        }

        private static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;

            }
            catch
            {
                return false;
            }
        }

        private static DataSet CheckDuplicateLog(string Docnum, string CompanyCode)
        {
            if (SAPCon12.State == ConnectionState.Closed) { SAPCon12.Open(); }
            DataSet dsetItem = new DataSet();
            SqlCommand CmdItem = new SqlCommand("select DocNum from syspex_ePo where DocNum ='" + Docnum + "' and Company ='" + CompanyCode + "'", SAPCon12)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SAPCon12.Close();
            return dsetItem;
        }

        private static DataSet GetPoNos(string CompanyCode)
        {
            SqlConnection SQLConnection = new SqlConnection();

            if (CompanyCode == "65ST")
                SQLConnection = SGConnection;
            SQLQuery = "select  Distinct top 5  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , T1.E_Mail, (CASE WHEN (select max( Country)  from CRD1 where  CardCode = (select CardCode from OPOR where DocNum = T0.DocNum)   and AdresType ='B') != 'SG' THEN 'alicia.ang@syspex.com,angela.yap@syspex.com' ELSE ''END) + ',' + T2.email + ',' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') + ',' + '65stproc@syspex.com,charlston.tong@syspex.com,wansing.teo@syspex.com' as [cc]  from OPOR T0 INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode  where  T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') and  CAST(T0.U_ePO AS nvarchar(max)) ='Yes' and T0.DocDate >='20200728' and  T0.DocDate <= getdate() and T0.DocStatus = 'O'";

            if (CompanyCode == "03SM")
            {
                //Go Live 20-06-17
                SQLConnection = KLConnection;
                SQLQuery = "select  Distinct top 5  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , T1.E_Mail,  T2.email + ',' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') + ',' + CASE WHEN T5.ItmsGrpCod = '101' THEN 'venice.tan@syspex.com,rabby.cheng@syspex.com' ELSE '' END as [cc]  from OPOR T0 INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode where year(t0.DocDate) = year(getdate()) and month(t0.DocDate) = month(getdate()) and day(t0.DocDate) = day(getdate()) and T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') and T0.DocDate >='20200617' and  T0.DocDate <= getdate() and T0.DocStatus = 'O'";
            }

            if (CompanyCode == "07ST")
            {
                //Go Live 22-06-20
                SQLConnection = JBConnection;
                SQLQuery = "select  Distinct top 5  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , T1.E_Mail,  T2.email + ',' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'')   as [cc]  from OPOR T0 INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode where year(t0.DocDate) = year(getdate()) and month(t0.DocDate) = month(getdate()) and day(t0.DocDate) = day(getdate()) and T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') and T0.DocDate >='20200622' and  T0.DocDate <= getdate()  and T0.DocStatus = 'O'";
            }

            if (CompanyCode == "21SK")
                SQLConnection = JKConnection;

            if (CompanyCode == "31SM")
                SQLConnection = SBConnection;

            if (CompanyCode == "04SI")
            {
                SQLConnection = PGConnection;
                SQLQuery = "select  Distinct top 5  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , T1.E_Mail,  T2.email + ',' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') +',' +'jane.khoo@syspex.com' + ',' + CASE WHEN T5.ItmsGrpCod = '101' THEN 'venice.tan@syspex.com,rabby.cheng@syspex.com' ELSE '' END as [cc]  from OPOR T0 INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode where year(t0.DocDate) = year(getdate()) and month(t0.DocDate) = month(getdate()) and day(t0.DocDate) = day(getdate()) and T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') and T0.DocDate >='20200617' and T0.DocDate <= getdate() and T0.DocStatus = 'O'";
            }
            DataSet dsetItem = new DataSet();
            SqlCommand CmdItem = new SqlCommand(SQLQuery, SQLConnection)
            {
                CommandType = CommandType.Text
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SQLConnection.Close();
            return dsetItem;
        }
        private static string ST_HTMLBULIDER(string DocNum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<p>Dear Supplier,</p>");
            sb.AppendLine("<p>Please find <strong><u>PO# " + DocNum + "</u></strong> and file attachments.</p>");
            sb.AppendLine("<p>We would need you:</p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li>To confirm the quantity and the delivery date within the next 24 hours <strong>(reply to all by using this e-mail</strong>)</li>");
            sb.AppendLine("<li>To indicate Syspex PO number for both Invoice and DO</li>");
            sb.AppendLine("<li> To take note our receiving hours (Monday to Fridays 8:00am &ndash; 12:00 &amp; 1:00pm &ndash; 4:00pm) <strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></li>");
            sb.AppendLine("<li> Please take note and comply that total height of incoming palletised goods should not exceed 1.5m. </ li>");
            sb.AppendLine("<li>***Do not reply to <u>noreply@syspex.com</u> as this is a unmanned mail box ***</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p style=\"color:black; background:yellow\">Please take note: We will be closed from 11 Feb to 15 Feb 2021 for CNY, our last receiving will be on 9 Feb 2021 (8:00am – 12:00 & 1:00pm – 4:00pm)</p>");
            sb.AppendLine("<p>Thank you for your co-operation.</p>");
            sb.AppendLine("<p>Best Regards,</p>");
            sb.AppendLine("<p>Syspex Procurement Team</p>");
            return sb.ToString();
        }
    }


}
