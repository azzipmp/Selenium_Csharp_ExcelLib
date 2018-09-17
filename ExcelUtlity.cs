using System;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;

using ExcelApp = Microsoft.Office.Interop.Excel; 

namespace TREDS.Virginia.Gov.QA.Utility.ExcelLib
{
   public class ExcelUtlity
    {
        private System.Data.OleDb.OleDbConnection MyConnection;
        private System.Data.DataSet DtSet;
        private System.Data.OleDb.OleDbDataAdapter MyCommand;
        private string _ConnectionString;
        public ExcelUtlity(String sExcelPath)
        {
            _ConnectionString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + sExcelPath + ";Extended Properties=Excel 8.0;";
        }

        public String getValue(String sSheetName, String sColName, int rownum)
        {
            string _value;
            try
            {
                using (MyConnection = new System.Data.OleDb.OleDbConnection(_ConnectionString))
                {
                    string sql = String.Format("select [{0}] from [{1}$] where id = {2}", sColName, sSheetName, rownum);
                    
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter(sql, MyConnection);

                    DtSet = new System.Data.DataSet();
                    MyCommand.FillSchema(DtSet, System.Data.SchemaType.Source);
                    MyCommand.Fill(DtSet);

                    _value = DtSet.Tables[0].Rows[0].ItemArray[0].ToString();
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("exception e" + ex);
                throw ex;
            }

            return _value;
        }
        public int rowCount(String sSheetName)
        {
            int count = 0;

            using (MyConnection = new System.Data.OleDb.OleDbConnection(_ConnectionString))
            {
                string sql = String.Format("select * from [" + sSheetName + "$]  ");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter(sql, MyConnection);

                DtSet = new System.Data.DataSet();
                MyCommand.FillSchema(DtSet, System.Data.SchemaType.Source);
                MyCommand.Fill(DtSet);

                count = DtSet.Tables[0].Rows.Count;
            }

            return count;
        }

        public Boolean writeValue(String sSheetName,int rowid ,String sColName, String value )
        {

            bool _value;

            try
            {

                using (MyConnection = new System.Data.OleDb.OleDbConnection(_ConnectionString))
                {
                    //string sql = String.Format("Update [" + sSheetName + "$] set [" + sColName + "] = '" + value + "' where id=" + rowid);
                    string sql = String.Format("Update [{0}$] set [{1}] = '{2}' where id = {4}", sSheetName, sColName, value, rowid);
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter(sql, MyConnection);

                    DtSet = new System.Data.DataSet();
                    MyCommand.FillSchema(DtSet, System.Data.SchemaType.Source);
                    MyCommand.Fill(DtSet);

                    _value = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("exception e" + ex);
                _value = false;
            }

            return _value;
        }

    }
}
