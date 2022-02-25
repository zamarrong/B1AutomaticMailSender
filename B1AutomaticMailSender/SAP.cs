using Sap.Data.Hana;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace B1AutomaticMailSender
{
    class SAP
    {
        public static SqlConnectionStringBuilder sql = new SqlConnectionStringBuilder();
        public static HanaConnectionStringBuilder hana = new HanaConnectionStringBuilder();
        public string CompanyName { get; set; }

        public SAP()
        {
            sql.DataSource = hana.Server = Properties.Settings.Default.DBServer;
            sql.InitialCatalog = hana.CurrentSchema = Properties.Settings.Default.DBName;
            sql.UserID = hana.UserName = Properties.Settings.Default.DBUserName;
            sql.Password = hana.Password = Properties.Settings.Default.DBPassword;
            sql.ConnectTimeout = hana.ConnectionTimeout = 30;

            CompanyName = GetCompanyName();
        }

        public class Document
        {
            public int DocEntry { get; set; }
            public string ObjType { get; set; }
            public int DocNum { get; set; }
            public DateTime DocDate { get; set; }
            public string DocCur { get; set; }
            public decimal DocRate { get; set; }
            public decimal DocTotal { get; set; }
            public string LicTradNum { get; set; }
            public string UUID { get; set; }

            public BusinessPartner BusinessPartner { get; set; }

            public Document()
            {
                BusinessPartner = new BusinessPartner();
            }

            public string DocTypeByObjType()
            {
                switch (ObjType)
                {
                    case "13":
                        return "IN";
                    case "14":
                        return "CM";
                    case "203":
                        return "DP";
                    default:
                        return "";
                }
            }

            public string MakePath()
            {
                string DateFolder = string.Format("{0}-{1}", DocDate.Year, DocDate.Month.ToString("00"));
                return string.Format(@"{0}\{1}\{2}", DateFolder, BusinessPartner.CardCode, DocTypeByObjType());
            }

        }

        public class BusinessPartner
        {
            public string CardCode { get; set; }
            public string CardName { get; set; }
            public string E_Mail { get; set; }

            public class ContactPerson
            {
                public string FirstName { get; set; }
                public string LastName { get; set; }
                public string E_Mail { get; set; }
            }

            public List<ContactPerson> GetContactPersons()
            {
                try
                {
                    string query = string.Format("SELECT OCPR.\"FirstName\", OCPR.\"LastName\", OCPR.\"E_MailL\" FROM OCPR WHERE OCPR.\"CardCode\" = '{0}' AND OCPR.\"NFeRcpn\" = 'Y'", CardCode);
                    DataTable contact_persons = ExecuteQuery(query);
                    List<ContactPerson> contact_persons_list = new List<ContactPerson>();
                    for (int i = 0; i < contact_persons.Rows.Count; i++)
                    {
                        ContactPerson contact_person = new ContactPerson();

                        contact_person.FirstName = (string)contact_persons.Rows[i]["FirstName"];
                        contact_person.LastName = (string)contact_persons.Rows[i]["LastName"];
                        contact_person.E_Mail = (string)contact_persons.Rows[i]["E_MailL"];

                        contact_persons_list.Add(contact_person);
                    }
                    return contact_persons_list;
                }
                catch (Exception ex)
                {
                    Program.WriteLog("GetContactPersons: " + ex.Message);
                    return new List<ContactPerson>();
                }
            }
        }

        public string GetCompanyName()
        {
            try
            {
                return ExecuteScalarString("SELECT \"CompnyName\" FROM OADM");
            }
            catch (Exception ex)
            {
                Program.WriteLog("GetCompanyName: " + ex.Message);
                return string.Empty;
            }
        }

        public List<Document> GetSendPendingDocuments(string ObjType, int Top)
        {
            try
            {
                string query = string.Format("SELECT TOP {1} {0}.\"DocEntry\", {0}.\"ObjType\", {0}.\"DocNum\", ECM2.\"ReportID\" \"UUID\", {0}.\"CardCode\", {0}.\"CardName\", OCRD.\"LicTradNum\", OCRD.\"E_Mail\", {0}.\"DocDate\", {0}.\"DocCur\", {0}.\"DocRate\", {0}.\"DocTotal\" FROM {0} INNER JOIN ECM2 ON ECM2.\"SrcObjAbs\" = {0}.\"DocEntry\" INNER JOIN OCRD ON OCRD.\"CardCode\" = {0}.\"CardCode\" WHERE {0}.\"U_EDocSent\" = 'N' AND ECM2.\"ReportID\" IS NOT NULL AND ECM2.\"SrcObjType\" = {2} AND OCRD.\"E_Mail\" IS NOT NULL", SAPTableByObjType(ObjType), Top, ObjType);
                DataTable documents = ExecuteQuery(query);
                List<Document> documents_list = new List<Document>();
                for (int i = 0; i < documents.Rows.Count; i++)
                {
                    Document document = new Document();

                    //Document
                    document.DocEntry = (int)documents.Rows[i]["DocEntry"];
                    document.ObjType = (string)documents.Rows[i]["ObjType"];
                    document.DocNum = (int)documents.Rows[i]["DocNum"];
                    document.DocDate = (DateTime)documents.Rows[i]["DocDate"];
                    document.DocCur = (string)documents.Rows[i]["DocCur"];
                    document.DocRate = (decimal)documents.Rows[i]["DocRate"];
                    document.DocTotal = (decimal)documents.Rows[i]["DocTotal"];

                    document.LicTradNum = (string)documents.Rows[i]["LicTradNum"];
                    document.UUID = (string)documents.Rows[i]["UUID"];

                    //Business Partner
                    document.BusinessPartner.CardCode = (string)documents.Rows[i]["CardCode"];
                    document.BusinessPartner.CardName = (string)documents.Rows[i]["CardName"];
                    document.BusinessPartner.E_Mail = (string)documents.Rows[i]["E_Mail"];

                    documents_list.Add(document);
                }
                return documents_list;
            }
            catch (Exception ex)
            {
                Program.WriteLog("GetSendPendingDocuments: " + ex.Message);
                return new List<Document>();
            }
        }
       
        public void UpdateEDocSentStatus(Document document, string status)
        {
            try
            {
                string query = string.Format("UPDATE {0} SET \"U_EDocSent\" = '{1}' WHERE \"DocEntry\" = {2}", SAPTableByObjType(document.ObjType), status, document.DocEntry);
                DataTable update = ExecuteQuery(query);
            }
            catch (Exception ex)
            {
                Program.WriteLog(ex.Message);
            }

        }

        public Document GetDocument(string UUID)
        {
            try
            {
                Document document = new Document();

                string ObjType = ExecuteScalarString(string.Format("SELECT TOP 1 \"SrcObjType\" FROM ECM2 WHERE \"ReportID\" = '{0}'", UUID));
                if (ObjType.Length > 0)
                {
                    string query = string.Format("SELECT TOP 1 {0}.\"DocEntry\", {0}.\"ObjType\", {0}.\"DocNum\", ECM2.\"ReportID\" \"UUID\", {0}.\"CardCode\", {0}.\"CardName\", OCRD.\"LicTradNum\", OCRD.\"E_Mail\", {0}.\"DocDate\", {0}.\"DocCur\", {0}.\"DocRate\", {0}.\"DocTotal\" FROM {0} INNER JOIN ECM2 ON ECM2.\"SrcObjAbs\" = {0}.\"DocEntry\" INNER JOIN OCRD ON OCRD.\"CardCode\" = {0}.\"CardCode\" WHERE ECM2.\"ReportID\"  = '{1}'", SAPTableByObjType(ObjType), UUID);
                    DataTable documents = ExecuteQuery(query);

                    //Document
                    document.DocEntry = (int)documents.Rows[0]["DocEntry"];
                    document.ObjType = (string)documents.Rows[0]["ObjType"];
                    document.DocNum = (int)documents.Rows[0]["DocNum"];
                    document.DocDate = (DateTime)documents.Rows[0]["DocDate"];
                    document.DocCur = (string)documents.Rows[0]["DocCur"];
                    document.DocRate = (decimal)documents.Rows[0]["DocRate"];
                    document.DocTotal = (decimal)documents.Rows[0]["DocTotal"];

                    document.LicTradNum = (string)documents.Rows[0]["LicTradNum"];
                    document.UUID = (string)documents.Rows[0]["UUID"];

                    //Business Partner
                    document.BusinessPartner.CardCode = (string)documents.Rows[0]["CardCode"];
                    document.BusinessPartner.CardName = (string)documents.Rows[0]["CardName"];
                    document.BusinessPartner.E_Mail = (string)documents.Rows[0]["E_Mail"];
                }

                return document;
            }
            catch (Exception ex)
            {
                Program.WriteLog("GetDocument: " + ex.Message);
                return new Document();
            }

        }
        public static DataTable ExecuteQuery(string query)
        {
            try
            {
                DB db = new DB();
                return (Properties.Settings.Default.HANA) ? db.ExecuteQueryHana(query) : db.ExecuteQuery(query);
            }
            catch (Exception ex)
            {
                Program.WriteLog(ex.Message);
                return null;
            }
        }

        public static string ExecuteScalarString(string query)
        {
            try
            {
                DB db = new DB();
                return (Properties.Settings.Default.HANA) ? db.ExecuteScalarStringHana(query) : db.ExecuteScalarString(query);
            }
            catch (Exception ex)
            {
                Program.WriteLog(ex.Message);
                return "";
            }
        }
        public static string SAPTableByObjType(string ObjType)
        {
            switch (ObjType)
            {
                case "13":
                    return "OINV";
                case "14":
                    return "ORIN";
                case "24":
                    return "ORCT";
                case "203":
                    return "ODPI";
                default:
                    return "";
            }
        }

        public static string DocNameByObjType(string ObjType)
        {
            switch (ObjType)
            {
                case "13":
                    return "Factura";
                case "14":
                    return "Nota de crédito";
                case "24":
                    return "Recepción de pagos";
                case "203":
                    return "Factura de anticipo";
                default:
                    return "";
            }
        }
    }
}
