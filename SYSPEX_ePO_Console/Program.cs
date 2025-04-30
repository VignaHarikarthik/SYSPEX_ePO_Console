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
            //  Go Live 10 / 08 / 2021

            EPO("65ST"); // Go Live on 31/08/20
            System.Threading.Thread.Sleep(100);
            EPO("03SM"); // Go Live on 17/06/20
            System.Threading.Thread.Sleep(100);
            EPO("07ST"); // Go Live on 17/06/20
            System.Threading.Thread.Sleep(100);

            EPO("21SK"); //Go Live on 20240311
            System.Threading.Thread.Sleep(100);
            EPO("31SM"); //Go Live on 20240311



        }


        public static bool export_pdf(string pdf_path, string db_name, string crystal_path, string docentry)
        {
            try
            {

                ReportDocument cryRpt = new ReportDocument();
                cryRpt.Load(crystal_path);

                new TableLogOnInfos();
                TableLogOnInfo crtableLogoninfo;
                var crConnectionInfo = new ConnectionInfo();

                ParameterFieldDefinitions crParameterFieldDefinitions;
                ParameterFieldDefinition crParameterFieldDefinition;
                ParameterValues crParameterValues = new ParameterValues();
                ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

                crParameterDiscreteValue.Value = Convert.ToString(docentry);
                crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
                crParameterFieldDefinition = crParameterFieldDefinitions["@DOCENTRY"];
                crParameterValues = crParameterFieldDefinition.CurrentValues;

                crParameterValues.Clear();
                crParameterValues.Add(crParameterDiscreteValue);
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

                crConnectionInfo.ServerName = "SYSPEXSAP04";
                crConnectionInfo.DatabaseName = db_name;
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

                CrDiskFileDestinationOptions.DiskFileName = pdf_path;
                CrExportOptions = cryRpt.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                cryRpt.Export();
                return true;


            }
            catch (CrystalReportsException ex)
            {

                throw ex;
            }
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
            // To = "vigna@syspex.com";

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

                if ((CompanyCode == "21SK"))
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_21SK.rpt");

                if ((CompanyCode == "31SM"))
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_31SM.rpt");

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


                if (CompanyCode == "65ST")
                    mm.Body = SG_HTMLBuilder(DocNum);

                if (CompanyCode == "04SI")
                    mm.Body = PG_HTMLBuilder(DocNum);

                if (CompanyCode == "03SM")
                    mm.Body = KL_HTMLBuilder(DocNum);

                if (CompanyCode == "07ST")
                    mm.Body = JB_HTMLBuilder(DocNum);

                if (CompanyCode == "21SK")
                    mm.Body = JK_HTMLBuilder(DocNum);

                if (CompanyCode == "31SM")
                    mm.Body = SB_HTMLBuilder(DocNum);


                //  mm.To.Add("vigna@syspex.com");



                foreach (var address in CC.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).Distinct())
                {
                    mm.CC.Add(new MailAddress(address)); //Adding Multiple CC email Id
                }


                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    EnableSsl = true
                };
                if (CompanyCode == "65ST")
                {
                    System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential("sg.procurement@syspex.com", "onzo yurc yoci zzzf"); //"enhance5");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                    smtp.Send(mm);
                }
                else if ((CompanyCode == "03SM") || (CompanyCode == "07ST"))
                {
                    System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential("procurement@syspex.com", "wacj rcui wqrt wcpa"); //"herself7");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                    smtp.Send(mm);
                }
                else if ((CompanyCode == "21SK") || (CompanyCode == "31SM"))
                {
                    System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential("noreply@syspex.com", "vani ktfs dhxq phcl");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                    smtp.Send(mm);
                }

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
            {
                SQLConnection = SGConnection;
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("SELECT DISTINCT TOP 5");
                sb.AppendLine("    T0.DocNum,");
                sb.AppendLine("    CONVERT(VARCHAR(10), T0.DocDate, 103) AS DocDate,");
                sb.AppendLine("    T3.DocEntry,");
                sb.AppendLine("    T0.CardCode,");
                sb.AppendLine("    T0.CardName,");
                sb.AppendLine("    T1.E_Mail,");
                sb.AppendLine("    (CASE");
                sb.AppendLine("        WHEN (SELECT MAX(Country) FROM CRD1");
                sb.AppendLine("              WHERE CardCode = (SELECT CardCode FROM OPOR WHERE DocNum = T0.DocNum)");
                sb.AppendLine("              AND AdresType = 'B') != 'SG'");
                sb.AppendLine("        THEN 'angela.yap@syspex.com,catherine.chang@syspex.com'");
                sb.AppendLine("        ELSE 'angela.yap@syspex.com'");
                sb.AppendLine("     END) + ',' + T2.Email + ',' +");
                sb.AppendLine("    -- PR Requester if multiple PRs exist in one PO");
                sb.AppendLine("    ISNULL((SELECT STUFF((SELECT ',' + Email");
                sb.AppendLine("                          FROM OHEM");
                sb.AppendLine("                          WHERE CAST(empid AS NVARCHAR) IN");
                sb.AppendLine("                              (SELECT Requester");
                sb.AppendLine("                               FROM OPRQ TX");
                sb.AppendLine("                               INNER JOIN PRQ1 TY ON TX.DocEntry = TY.DocEntry");
                sb.AppendLine("                               WHERE TY.TrgetEntry = T0.DocEntry)");
                sb.AppendLine("                          FOR XML PATH('')), 1, 1, '')), '') + ',' +");
                sb.AppendLine("    -- PR Owner Code if multiple PRs exist in one PO");
                sb.AppendLine("    ISNULL((SELECT STUFF((SELECT ',' + Email");
                sb.AppendLine("                          FROM OHEM");
                sb.AppendLine("                          WHERE CAST(empid AS NVARCHAR) IN");
                sb.AppendLine("                              (SELECT TX.OwnerCode");
                sb.AppendLine("                               FROM OPRQ TX");
                sb.AppendLine("                               INNER JOIN PRQ1 TY ON TX.DocEntry = TY.DocEntry");
                sb.AppendLine("                               WHERE TY.TrgetEntry = T0.DocEntry)");
                sb.AppendLine("                          FOR XML PATH('')), 1, 1, '')), '') + ',' +");
                sb.AppendLine("    'procurement@syspex.com' + ',' + 'SG.Procurement@syspex.com' AS [cc]");
                sb.AppendLine("FROM OPOR T0");
                sb.AppendLine("INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode");
                sb.AppendLine("INNER JOIN OHEM T2 ON T2.empID = T0.OwnerCode");
                sb.AppendLine("INNER JOIN POR1 T3 ON T3.DocEntry = T0.DocEntry");
                sb.AppendLine("WHERE T0.DocNum NOT IN");
                sb.AppendLine("    (SELECT DocNum FROM [AndriodAppDB].[dbo].[syspex_ePO]");
                sb.AppendLine("     WHERE Company = '" + CompanyCode + "')");
                sb.AppendLine("AND CAST(T0.U_ePO AS NVARCHAR(MAX)) = 'Yes'");
                sb.AppendLine("AND T0.DocDate <= GETDATE()");
                sb.AppendLine("AND T0.DocStatus = 'O'");

                SQLQuery = sb.ToString();

                //changes
            }

            if (CompanyCode == "03SM")
            {
                //Go Live 20-06-17
                SQLConnection = KLConnection;

                SQLQuery = "select  Distinct top 8  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , " +
                    "T1.E_Mail,  T2.email + ',' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' + " +
                    "ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') + ',jane.khoo@syspex.com,procurement@syspex.com,peyyin.lim@syspex.com,vigna@syspex.com,esther.lim@syspex.com' +',' +" +
                    " CASE WHEN T5.SlpName like  'TE01%'  THEN 'peyyin.lim@syspex.com,vigna@syspex.com' ELSE '' END as [cc] " +
                    " from OPOR T0 INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 " +
                    "T3 on T3.DocEntry = T0.DocEntry INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry " +
                    "INNER JOIN OSLP T5 on T5.SlpCode = T0.SlpCode where year(t0.DocDate) = year(getdate()) and month(t0.DocDate) = month(getdate())" +
                    " and T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') and T0.DocDate >='20250101' and CAST(T0.U_equote AS nvarchar(max)) ='Yes' and T0.DocDate <= getdate() and T0.DocStatus = 'O'";
            }

            if (CompanyCode == "07ST")
            {
                //Go Live 22-06-20
                SQLConnection = JBConnection;
                SQLQuery = "select  Distinct top 5  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , T1.E_Mail,  T2.email + ',' " +
                    "+ ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' +'procurement@syspex.com,peyyin.lim@syspex.com,' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') + ',' + " +
                    "CASE WHEN T5.SlpName like  'TE01%' THEN 'peyyin.lim@syspex.com,ivy.lim@syspex.com' ELSE '' END   as [cc]  from OPOR T0 INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry" +
                    " INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry INNER JOIN OSLP T5 on T5.SlpCode = T0.SlpCode where year(t0.DocDate) = year(getdate()) and month(t0.DocDate) = month(getdate())  and T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') " +
                    "and T0.DocDate >='20250101' and CAST(T0.U_equote AS nvarchar(max)) ='Yes' and  T0.DocDate <= getdate()  and T0.DocStatus = 'O'";
            }

            if (CompanyCode == "21SK")
            {
                SQLConnection = JKConnection;

                SQLQuery = @"select 
                                Distinct top 5 
                                T0.DocNum,
                                CONVERT(VARCHAR(10), T0.DocDate, 103) as DocDate, 
                                T3.DocEntry,
                                T0.CardCode, 
                                T0.CardName, 
                                T1.E_Mail,
                                '21skprocurement@syspex.com,procurement@syspex.com,' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') + ',' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') as [Cc]  
                             from OPOR T0 
                             INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode
                             INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode
                             INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry
                             INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry 
                             INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode 
                             where year(T0.DocDate) = year(getdate()) 
                             and month(T0.DocDate) = month(getdate())
                             and T0.DocNum not in (select DocNum from [AndriodAppDB].[dbo].[syspex_ePO] where Company = '21SK') 
                             and T0.DocDate >= '20240311' 
                             and CAST(T0.U_equote AS nvarchar(max)) = 'Yes' 
                             and T0.DocDate <= getdate() 
                             and T0.DocStatus = 'O'
                             order by T0.DocNum";
            }

            if (CompanyCode == "31SM")
            {
                SQLConnection = SBConnection;

                SQLQuery = @"select 
                                Distinct top 5 
                                T0.DocNum,
                                CONVERT(VARCHAR(10), T0.DocDate, 103) as DocDate, 
                                T3.DocEntry,
                                T0.CardCode, 
                                T0.CardName, 
                                T1.E_Mail,
                                '21skprocurement@syspex.com,procurement@syspex.com,' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') + ',' + ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') as [Cc]  
                             from OPOR T0 
                             INNER JOIN OCRD T1 on T0.CardCode = T1.CardCode
                             INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode
                             INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry
                             INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry 
                             INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode 
                             where year(T0.DocDate) = year(getdate()) 
                             and month(T0.DocDate) = month(getdate())
                             and T0.DocNum not in (select DocNum from [AndriodAppDB].[dbo].[syspex_ePO] where Company = '31SM') 
                             and T0.DocDate >= '20240311' 
                             and CAST(T0.U_equote AS nvarchar(max)) = 'Yes' 
                             and T0.DocDate <= getdate() 
                             and T0.DocStatus = 'O'
                             order by T0.DocNum";
            }



            if (CompanyCode == "04SI")
            {
                SQLConnection = PGConnection;
                SQLQuery = "select  Distinct top 5  T0.DocNum,CONVERT(VARCHAR(10), T0.docdate, 103) as docdate, T3.DocEntry,T0.CardCode, T0.CardName , T1.E_Mail,  T2.email + ',' + ISNULL((SELECT Email from OHEM where CAST(empid as nvarchar) = CAST(T4.Requester as nvarchar)),'') +',' " +
                    "+ ISNULL((SELECT Email from OHEM where empid = T4.OwnerCode ),'') +',' +',procurement@syspex.com,jane.khoo@syspex.com' + ',' + CASE WHEN T5.ItmsGrpCod = '101' THEN 'venice.tan@syspex.com,peyyin.lim@syspex.com' ELSE '' END as [cc]  from OPOR T0 INNER JOIN OCRD T1 " +
                    "on T0.CardCode = T1.CardCode INNER JOIN OHEM T2 on T2.empID = T0.OwnerCode INNER JOIN POR1 T3 on T3.DocEntry = T0.DocEntry INNER JOIN OPRQ T4 on T4.DocEntry = T3.BaseEntry INNER JOIN OITM T5 on T5.ItemCode = T3.ItemCode where year(t0.DocDate) = year(getdate()) and month(t0.DocDate) = month(getdate())" +
                    " and T0.DocNum not in (select DocNum from[AndriodAppDB].[dbo].[syspex_ePO] where Company = '" + CompanyCode + "') and T0.DocDate >='20200617' and CAST(T0.U_equote AS nvarchar(max)) ='Yes' and T0.DocDate <= getdate() and T0.DocStatus = 'O'";
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

        private static string SG_HTMLBuilder(string DocNum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Dear Supplier,</p>");
            sb.AppendLine("<p>Please find <strong><u>PO# " + DocNum + "</u></strong> and file attachments.</p>");
            sb.AppendLine("<p>Reply back this email to confirm on the order quantity and the delivery date stated on the PO within the next 24 hours</p>");
            sb.AppendLine("<p>Kindly take note and comply with the following packaging and delivery information, </p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li>To indicate Syspex PO number for both Invoice and DO.</li>");
            sb.AppendLine("<li>To indicate serial number on each outer packaging (When applicable).</li>");
            sb.AppendLine("<li> To take note our receiving hours (Monday to Fridays 10:00am &ndash; 12:00 &amp; 1:00pm &ndash; 4:00pm).<strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></li>");
            sb.AppendLine("<li> Please take note and comply that total height of incoming palletised goods should not exceed 1.5m.</ li>");
            sb.AppendLine("<li> The pallet must be able to truck by hand pallet truck.</li>");
            sb.AppendLine("<li> Please email us soft copy of invoice and packing list once shipment ready for dispatch.</li>");
            sb.AppendLine("<li> For multiple package shipment, please indicate content list on outside of each package.</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p>Thank you for your co-operation.</p>");
            sb.AppendLine("<p>Best Regards,</p>");
            sb.AppendLine("<p>Syspex Procurement Team</p>");
            return sb.ToString();
        }
        private static string KL_HTMLBuilder(string docnum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Dear Supplier,</p>");
            sb.AppendLine("<p>Please find <strong><u>PO# " + docnum + "</u></strong>&nbsp;and file attachments.</p>");
            sb.AppendLine("<p><strong><u>Please acknowledge this email</u></strong> to <strong><u>confirm on the</u></strong> <strong><u>order quantity and the delivery date</u></strong> stated on the PO <strong><u>within the next 24 hours</u></strong></p>");
            sb.AppendLine("<p>Kindly take note and comply with the following packaging and delivery information,</p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li>To indicate Syspex PO number for both Invoice and DO.</li>");
            sb.AppendLine("<li>To indicate item description &amp; serial number on each outer packaging (When applicable).</li>");
            sb.AppendLine("<li>To take note our receiving hours as below:</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("    <p align=\"center\">Monday to Thursday: 8:30am &ndash; 12:00pm &amp; 1:00pm &ndash; 5:30pm</p>");
            sb.AppendLine("<p align=\"center\">Friday: 8.30am &ndash; 1:00pm &amp; 2:30pm &ndash; 5:30pm</p>");
            sb.AppendLine("<p><strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></p>");
            sb.AppendLine("<ol start=4>");
            sb.AppendLine("<li>The pallet must be able to truck by hand pallet truck.</li>");
            sb.AppendLine("<li>Please email us soft copy of invoice and packing list once shipment ready for dispatch.</li>");
            sb.AppendLine("<li>For multiple package shipment, please indicate item description on outside of each package.</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p>Thank you for your co-operation.</p>");
            sb.AppendLine("<p>Best Regards,</p>");
            sb.AppendLine("<p>Syspex Purchasing Team</p>");


            return sb.ToString();
        }
        private static string JB_HTMLBuilder(string docnum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Dear Supplier,</p>");
            sb.AppendLine("<p>Please find <strong><u>PO# " + docnum + "</u></strong> and file attachments.</p>");
            sb.AppendLine("<p><strong><u>Please acknowledge this email</u></strong> to <strong><u>confirm on the</u></strong><u> <strong>order quantity and the delivery date</strong></u> stated on the PO <strong><u>within the next 24 hours</u></strong></p>");
            sb.AppendLine("<p>Kindly take note and comply with the following packaging and delivery information,</p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li>To indicate Syspex PO number for both Invoice and DO.</li>");
            sb.AppendLine("<li>To indicate item description &amp; serial number on each outer packaging (When applicable).</li>");
            sb.AppendLine("<li>To take note our receiving hours as below:</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p align=\"center\">Monday to Thursday: 8:00am &ndash; 12:30pm &amp; 1:30pm &ndash; 5:00pm</p>");
            sb.AppendLine("<p align=\"center\">Friday: 8.00am &ndash; 12:30pm &amp; 2:30pm &ndash; 5:00pm</p>");
            sb.AppendLine("<p><strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></p>");
            sb.AppendLine("<ol start=4>");
            sb.AppendLine("<li>The pallet must be able to truck by hand pallet truck.</li>");
            sb.AppendLine("<li>Please email us soft copy of invoice and packing list once shipment ready for dispatch.</li>");
            sb.AppendLine("<li>For multiple package shipment, please indicate item description on outside of each package.</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p>Thank you for your co-operation.</p>");
            sb.AppendLine("<p>Best Regards,</p>");
            sb.AppendLine("<p>Syspex Purchasing Team</p>");

            return sb.ToString();
        }

        private static string PG_HTMLBuilder(string docnum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p><span>Dear Supplier,</span></p>");
            sb.AppendLine("<p><span>Please find<span>&nbsp;</span><strong><u>PO# " + docnum + "</u></strong>&nbsp;and file attachments.</span></p>");
            sb.AppendLine("<p><strong><u><span>Please acknowledge this email</span></u></strong><span><span>&nbsp;</span>to<span>&nbsp;</span><strong><u>confirm on the</u></strong><u><span>&nbsp;</span><strong>order quantity and the delivery date</strong></u><span>&nbsp;</span>stated on the PO<span>&nbsp;</span><strong><u>within the next 24 hours</u></strong></span></p>");
            sb.AppendLine("<p><span>Kindly take note and comply with the following packaging and delivery information,</span></p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li><span>To indicate Syspex PO number for both Invoice and DO.</span></li>");
            sb.AppendLine("<li><span>To indicate item description &amp; serial number on each outer packaging (When applicable).</span></li>");
            sb.AppendLine("<li><span>To take note our receiving hours as below:</span></li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p align=\"center\"><span>Monday to Thursday: 8:00am &ndash; 12:00pm &amp; 1:00pm &ndash; 5:00pm</span></p>");
            sb.AppendLine("<p align=\"center\"><span>Friday: 8.00am &ndash; 12:00pm &amp; 2:00pm &ndash; 5:00pm</span></p>");
            sb.AppendLine("<p ><strong><span>- Only applicable to supplier(s) deliver at Syspex Warehouse</span></strong></p>");
            sb.AppendLine("<ol start=4>");
            sb.AppendLine("<li><span>The pallet must be able to truck by hand pallet truck.</span></li>");
            sb.AppendLine("<li><span>Please email us soft copy of invoice and packing list once shipment ready for dispatch.</span></li>");
            sb.AppendLine("<li><span>For multiple package shipment, please indicate item description on outside of each package.</span></li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p ><span>Thank you for your co-operation.</span></p>");
            sb.AppendLine("<p><span>Best Regards,</span></p>");
            sb.AppendLine("<p ><span>Syspex Purchasing Team</span></p>");

            return sb.ToString();
        }
        private static string JK_HTMLBuilder(string docnum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p><strong><i>This email is made with dual language support (English - upper section and Indonesian - lower section) as can be seen on below</i></strong></p>");
            sb.AppendLine("<p style=\"color:red;\"><strong><i>English version</i></strong></p>");
            sb.AppendLine("<p><span>Dear Supplier,</span></p>");
            sb.AppendLine("<p><span>Please find<span>&nbsp;</span><strong><u>PO# " + docnum + "</u></strong>&nbsp;and file attachment(s).</span></p>");
            sb.AppendLine("<p><span>Reply back this email to confirm on the order quantity and the delivery date stated on the PO within the next 24 hours.</span></p>");
            sb.AppendLine("<p><span>Kindly take note and comply with the following packaging and delivery information.</span></p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li><span>To indicate Syspex PO number for both Invoice and DO.</span></li>");
            sb.AppendLine("<li><span>To indicate serial number on each outer packaging (When applicable).</span></li>");
            sb.AppendLine("<li><span>To take note our receiving hours (Monday to Fridays 08:00am – 12:00 & 1:00pm – 4:00pm). <strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></span></li>");
            sb.AppendLine("<li><span>Please take note and comply that total height of incoming palletised goods should not exceed 2m.</span></li>");
            sb.AppendLine("<li><span>The pallet must be able to truck by hand pallet truck.</span></li>");
            sb.AppendLine("<li><span>Please email us soft copy of invoice and packing list once shipment ready for dispatch.</span></li>");
            sb.AppendLine("<li><span>For multiple package shipment, please indicate content list on outside of each package.</span></li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p><span>Thank you for your co-operation.</span></p>");
            sb.AppendLine("<p><span>Best Regards,</span></p>");
            sb.AppendLine("<p><span>Syspex Procurement Team</span></p>");
            sb.AppendLine("<br><br>");
            sb.AppendLine("<p style=\"color:red;\"><strong><i>Indonesian version</i></strong></p>");
            sb.AppendLine("<p><span>Kepada Yth. Rekanan Syspex,</span></p>");
            sb.AppendLine("<p><span>Terlampir adalah<span>&nbsp;</span><strong><u>PO# " + docnum + "</u></strong>&nbsp;dan lampiran yang diperlukan.</span></p>");
            sb.AppendLine("<p><span>Mohon tolong balas kembali email ini untuk mengkonfirmasi jumlah dan tanggal pengiriman yang tertera pada PO dalam waktu 24 jam ke depan.</span></p>");
            sb.AppendLine("<p><span>Harap perhatikan dan patuhi informasi pengemasan dan pengiriman berikut.</span></p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li><span>Untuk menunjukkan nomor PO Syspex untuk Invoice dan DO.</span></li>");
            sb.AppendLine("<li><span>Untuk menunjukkan nomor seri pada setiap kemasan luar (Bila ada).</span></li>");
            sb.AppendLine("<li><span>Perhatikan jam penerimaan kami (Senin sampai Jumat 08:00 – 12:00 & 13:00 – 16:00). <strong>- Hanya berlaku untuk pengiriman di Gudang Syspex</strong></span></li>");
            sb.AppendLine("<li><span>Harap perhatikan dan patuhi bahwa tinggi total barang palet yang masuk tidak boleh melebihi 2m.</span></li>");
            sb.AppendLine("<li><span>Palet harus dapat diangkut dengan truk palet tangan.</span></li>");
            sb.AppendLine("<li><span>Silakan kirim email kepada kami salinan faktur dan daftar pengepakan setelah pengiriman siap untuk dikirim.</span></li>");
            sb.AppendLine("<li><span>Untuk pengiriman beberapa paket, harap cantumkan daftar konten di luar setiap paket.</span></li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p><span>Terima kasih atas kerja sama anda.</span></p>");
            sb.AppendLine("<p><span>Salam,</span></p>");
            sb.AppendLine("<p><span>Tim Pengadaan Syspex</span></p>");
            return sb.ToString();
        }

        private static string SB_HTMLBuilder(string docnum)
        {
            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p><strong><i>This email is made with dual language support (English - upper section and Indonesian - lower section) as can be seen on below</i></strong></p>");
            sb.AppendLine("<p style=\"color:red;\"><strong><i>English version</i></strong></p>");
            sb.AppendLine("<p><span>Dear Supplier,</span></p>");
            sb.AppendLine("<p><span>Please find<span>&nbsp;</span><strong><u>PO# " + docnum + "</u></strong>&nbsp;and file attachment(s).</span></p>");
            sb.AppendLine("<p><span>Reply back this email to confirm on the order quantity and the delivery date stated on the PO within the next 24 hours.</span></p>");
            sb.AppendLine("<p><span>Kindly take note and comply with the following packaging and delivery information.</span></p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li><span>To indicate Syspex PO number for both Invoice and DO.</span></li>");
            sb.AppendLine("<li><span>To indicate serial number on each outer packaging (When applicable).</span></li>");
            sb.AppendLine("<li><span>To take note our receiving hours (Monday to Fridays 08:00am – 12:00 & 1:00pm – 4:00pm). <strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></span></li>");
            sb.AppendLine("<li><span>Please take note and comply that total height of incoming palletised goods should not exceed 2m.</span></li>");
            sb.AppendLine("<li><span>The pallet must be able to truck by hand pallet truck.</span></li>");
            sb.AppendLine("<li><span>Please email us soft copy of invoice and packing list once shipment ready for dispatch.</span></li>");
            sb.AppendLine("<li><span>For multiple package shipment, please indicate content list on outside of each package.</span></li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p><span>Thank you for your co-operation.</span></p>");
            sb.AppendLine("<p><span>Best Regards,</span></p>");
            sb.AppendLine("<p><span>Syspex Procurement Team</span></p>");
            sb.AppendLine("<br><br>");
            sb.AppendLine("<p style=\"color:red;\"><strong><i>Indonesian version</i></strong></p>");
            sb.AppendLine("<p><span>Kepada Yth. Rekanan Syspex,</span></p>");
            sb.AppendLine("<p><span>Terlampir adalah<span>&nbsp;</span><strong><u>PO# " + docnum + "</u></strong>&nbsp;dan lampiran yang diperlukan.</span></p>");
            sb.AppendLine("<p><span>Mohon tolong balas kembali email ini untuk mengkonfirmasi jumlah dan tanggal pengiriman yang tertera pada PO dalam waktu 24 jam ke depan.</span></p>");
            sb.AppendLine("<p><span>Harap perhatikan dan patuhi informasi pengemasan dan pengiriman berikut.</span></p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li><span>Untuk menunjukkan nomor PO Syspex untuk Invoice dan DO.</span></li>");
            sb.AppendLine("<li><span>Untuk menunjukkan nomor seri pada setiap kemasan luar (Bila ada).</span></li>");
            sb.AppendLine("<li><span>Perhatikan jam penerimaan kami (Senin sampai Jumat 08:00 – 12:00 & 13:00 – 16:00). <strong>- Hanya berlaku untuk pengiriman di Gudang Syspex</strong></span></li>");
            sb.AppendLine("<li><span>Harap perhatikan dan patuhi bahwa tinggi total barang palet yang masuk tidak boleh melebihi 2m.</span></li>");
            sb.AppendLine("<li><span>Palet harus dapat diangkut dengan truk palet tangan.</span></li>");
            sb.AppendLine("<li><span>Silakan kirim email kepada kami salinan faktur dan daftar pengepakan setelah pengiriman siap untuk dikirim.</span></li>");
            sb.AppendLine("<li><span>Untuk pengiriman beberapa paket, harap cantumkan daftar konten di luar setiap paket.</span></li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p><span>Terima kasih atas kerja sama anda.</span></p>");
            sb.AppendLine("<p><span>Salam,</span></p>");
            sb.AppendLine("<p><span>Tim Pengadaan Syspex</span></p>");
            return sb.ToString();
        }

    }


}
