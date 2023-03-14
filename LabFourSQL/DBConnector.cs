using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using System.Data;

namespace LabFourSQL
{
    class DBConnector
    {
        private static string connectStr = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\#\\LabFourSQL\\LabFourSQL\\CompanyDB.mdf;Integrated Security=True";

        public static void WriteDataDB(List<Company> inputArray) {
            using (SqlConnection openCon = new SqlConnection(connectStr)) {

                string saveCompany = "INSERT into CompanyTable (Bankrupt, AccountsReceivableTurnover, OperatingGrossMargin, CashCurrentLiability) VALUES (@bankrupt, @accountsReceivableTurnover, @operatingGrossMargin, @cashCurrentLiability)";

                openCon.Open();

                foreach (Company company in inputArray) {
                    using (SqlCommand querySaveCompany = new SqlCommand(saveCompany)) {
                        querySaveCompany.Connection = openCon;

                        querySaveCompany.Parameters.Add("@bankrupt", SqlDbType.Bit).Value = company.getBankrupt();
                        querySaveCompany.Parameters.Add("@accountsReceivableTurnover", SqlDbType.Float).Value = company.getAccountsReceivableTurnover();
                        querySaveCompany.Parameters.Add("@operatingGrossMargin", SqlDbType.Float).Value = company.getOperatingGrossMargin();
                        querySaveCompany.Parameters.Add("@cashCurrentLiability", SqlDbType.Float).Value = company.getCashCurrentLiability();
                        querySaveCompany.ExecuteNonQuery();

                    }
                }
                openCon.Close();
            }
            
        }
        public static List<Company> ReadDataFromDB()
        {
            List<Company> outputArray = new List<Company>();
            using (SqlConnection openCon = new SqlConnection(connectStr))
            {
                string readCompany = "SELECT Bankrupt, AccountsReceivableTurnover, OperatingGrossMargin, CashCurrentLiability FROM CompanyTable";
                openCon.Open();

                using (SqlCommand queryReadCompany = new SqlCommand(readCompany))
                {
                    queryReadCompany.Connection = openCon;
                    SqlDataReader reader = queryReadCompany.ExecuteReader();    // вся таблица
                    while (reader.Read())       //считывание отдельной строки и переход к следующей
                    { 
                        outputArray.Add(new Company(Convert.ToBoolean(reader[0]), Convert.ToDouble(reader[1]), Convert.ToDouble(reader[2]), Convert.ToDouble(reader[3])));
                    }
                    reader.Close();
                }
                openCon.Close();
            }
            return outputArray;

        }

    }
}
