using System;
using System.Data.SqlServerCe;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class SQLiteHelper_Old
	{
		public static string strDbFile = "gwbl.gwbldb";

		public static string strGwgsDbFile = "C:\\Program Files (x86)\\公文助手\\gwbl.gwbldb";

		public SqlCeConnection dbConnection;

		private SqlCeCommand dbCommand;

		private SqlCeDataReader dataReader;

		private SqlCeConnectionStringBuilder dbConnectionstr;

		public SQLiteHelper_Old(string connectionString)
		{
			try
			{
				dbConnection = new SqlCeConnection();
				dbConnectionstr = new SqlCeConnectionStringBuilder();
				dbConnectionstr.DataSource = connectionString;
				dbConnectionstr.Password = "denvy1987";
				dbConnection.ConnectionString = dbConnectionstr.ToString();
				dbConnection.Open();
			}
			catch (Exception ex)
			{
				Log(ex.ToString());
			}
		}

		public SQLiteHelper_Old()
		{
			try
			{
				dbConnection = new SqlCeConnection();
				dbConnectionstr = new SqlCeConnectionStringBuilder();
				string text = CommonConfig.strBaseFolder + "\\gwgs.sdf";
				text = text.Replace("公文助手插件", "公文助手");
				dbConnectionstr.DataSource = text;
				dbConnectionstr.Password = "denvy1987";
				dbConnection.ConnectionString = dbConnectionstr.ToString();
				dbConnection.Open();
			}
			catch (Exception ex)
			{
				Log(ex.ToString());
				MessageBox.Show(ex.ToString());
			}
		}

		public SqlCeDataReader ExecuteQuery(string queryString)
		{
			try
			{
				dbCommand = dbConnection.CreateCommand();
				dbCommand.CommandText = queryString;
				dataReader = dbCommand.ExecuteReader();
			}
			catch (Exception ex)
			{
				Log(ex.Message);
			}
			return dataReader;
		}

		public void CloseConnection()
		{
			if (dbCommand != null)
			{
				dbCommand.Cancel();
			}
			dbCommand = null;
			if (dataReader != null)
			{
				dataReader.Close();
			}
			dataReader = null;
			if (dbConnection != null)
			{
				dbConnection.Close();
			}
			dbConnection = null;
		}

		public SqlCeDataReader ReadFullTable(string tableName)
		{
			string queryString = "SELECT * FROM " + tableName;
			return ExecuteQuery(queryString);
		}

		public SqlCeDataReader InsertValues(string tableName, string[] values)
		{
			int fieldCount = ReadFullTable(tableName).FieldCount;
			int num = values.Length;
			string text = "INSERT INTO " + tableName + " VALUES ('" + values[0] + "'";
			for (int i = 1; i < values.Length; i++)
			{
				text = text + ", '" + values[i] + "'";
			}
			text += " )";
			return ExecuteQuery(text);
		}

		public SqlCeDataReader InsertValues_WithC(string tableName, string colName, string[] values)
		{
			string text = "INSERT INTO " + tableName + "(" + colName + ") VALUES ('" + values[0] + "'";
			for (int i = 1; i < values.Length; i++)
			{
				text = text + ", '" + values[i] + "'";
			}
			text += " )";
			return ExecuteQuery(text);
		}

		public SqlCeDataReader UpdateValues(string tableName, string[] colNames, string[] colValues, string key, string value, string operation)
		{
			int num = colNames.Length;
			int num2 = colValues.Length;
			string text = "UPDATE " + tableName + " SET " + colNames[0] + "='" + colValues[0] + "'";
			for (int i = 1; i < colValues.Length; i++)
			{
				text = text + ", " + colNames[i] + "='" + colValues[i] + "'";
			}
			text = text + " WHERE " + key + operation + "'" + value + "'";
			return ExecuteQuery(text);
		}

		public SqlCeDataReader UpdateValues(string tableName, string[] colNames, string[] colValues, string key1, string value1, string operation, string key2, string value2)
		{
			int num = colNames.Length;
			int num2 = colValues.Length;
			string text = "UPDATE " + tableName + " SET " + colNames[0] + "='" + colValues[0] + "'";
			for (int i = 1; i < colValues.Length; i++)
			{
				text = text + ", " + colNames[i] + "='" + colValues[i] + "'";
			}
			text = text + " WHERE " + key1 + operation + "'" + value1 + "'OR " + key2 + operation + "'" + value2 + "'";
			return ExecuteQuery(text);
		}

		public SqlCeDataReader DeleteValuesOR(string tableName, string[] colNames, string[] colValues, string[] operations)
		{
			if (colNames.Length == colValues.Length && operations.Length == colNames.Length)
			{
				int num = operations.Length;
				int num2 = colValues.Length;
			}
			string text = "DELETE FROM " + tableName + " WHERE " + colNames[0] + operations[0] + "'" + colValues[0] + "'";
			for (int i = 1; i < colValues.Length; i++)
			{
				text = text + "OR " + colNames[i] + operations[0] + "'" + colValues[i] + "'";
			}
			return ExecuteQuery(text);
		}

		public SqlCeDataReader DeleteValuesAND(string tableName, string[] colNames, string[] colValues, string[] operations)
		{
			if (colNames.Length == colValues.Length && operations.Length == colNames.Length)
			{
				int num = operations.Length;
				int num2 = colValues.Length;
			}
			string text = "DELETE FROM " + tableName + " WHERE " + colNames[0] + operations[0] + "'" + colValues[0] + "'";
			for (int i = 1; i < colValues.Length; i++)
			{
				text = text + " AND " + colNames[i] + operations[i] + "'" + colValues[i] + "'";
			}
			return ExecuteQuery(text);
		}

		public SqlCeDataReader CreateTable(string tableName, string[] colNames, string[] colTypes)
		{
			string text = "CREATE TABLE IF NOT EXISTS " + tableName + "( " + colNames[0] + " " + colTypes[0];
			for (int i = 1; i < colNames.Length; i++)
			{
				text = text + ", " + colNames[i] + " " + colTypes[i];
			}
			text += "  ) ";
			return ExecuteQuery(text);
		}

		public SqlCeDataReader ReadTable(string tableName, string[] items, string[] colNames, string[] operations, string[] colValues)
		{
			string text = "SELECT " + items[0];
			for (int i = 1; i < items.Length; i++)
			{
				text = text + ", " + items[i];
			}
			text = text + " FROM " + tableName + " WHERE " + colNames[0] + " " + operations[0] + " " + colValues[0];
			for (int j = 0; j < colNames.Length; j++)
			{
				text = text + " AND " + colNames[j] + " " + operations[j] + " " + colValues[0] + " ";
			}
			return ExecuteQuery(text);
		}

		public void VACUUM()
		{
			ExecuteQuery("VACUUM");
		}

		private static void Log(string s)
		{
			Console.WriteLine("class SQLiteHelper:::" + s);
		}
	}
}
