using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using static System.Configuration.ConfigurationSettings;

namespace CorrectContraAccountLogicDLL
{
    [System.Obsolete("This method is obsolete, it has been replaced by System.Configuration!System.Configuration.ConfigurationManager.AppSettings")]
    public class DiConnectionCompany : ICompany
    {
        public DiConnectionCompany()
        {
            Company = new CompanyClass
            {
                Server = AppSettings["Server"],
                DbServerType = BoDataServerTypes.dst_MSSQL2016,
                UserName = AppSettings["UserName"],
                Password = AppSettings["Password"],
                CompanyDB = AppSettings["CompanyDB"],
                language = BoSuppLangs.ln_English
            };
            Company.Connect();
            if (!Company.Connected)
            {
                throw new Exception($"Cannot Connect To the Server :  {Company.Server}, {Company.UserName}, {Company.CompanyDB}");
            }
        }

        public DiConnectionCompany(string server, int dbServerType, string userName, string password, string companyDb)
        {
            Company = new CompanyClass
            {
                Server = server,
                DbServerType = BoDataServerTypes.dst_MSSQL2016,
                UserName = userName,
                Password = password,
                CompanyDB = companyDb,
                language = BoSuppLangs.ln_English
            };
            Company.Connect();
            if (!Company.Connected)
            {
                throw new Exception($"Cannot Connect To the Server :  {Company.Server}, {Company.UserName}, {Company.CompanyDB}");
            }
        }

        public Company Company { get; set; }
    }
}

