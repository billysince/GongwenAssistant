using System;
using System.Data.SqlServerCe;
using System.Windows.Forms;

namespace Local_Wps_Vsto
{
	public class SQLiteHelper
	{
		public const string strDbFile = "gwgs.sdf";

		public static string strGwgsDbFile = "C:\\Program Files (x86)\\公文助手\\gwbl.gwbldb";

		public SqlCeConnection dbConnection;

		private SqlCeCommand dbCommand;

		private SqlCeDataReader dataReader;

		private SqlCeConnectionStringBuilder dbConnectionstr;

		public SQLiteHelper(string connectionString)
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

		public SQLiteHelper()
		{
			string dataSource = CommonConfig.strBaseFolder_Common + "\\gwgs.sdf";
			try
			{
				dbConnection = new SqlCeConnection();
				dbConnectionstr = new SqlCeConnectionStringBuilder();
				dbConnectionstr.DataSource = dataSource;
				dbConnectionstr.Password = "Denvy1987";
				dbConnection.ConnectionString = dbConnectionstr.ToString();
				dbConnection.Open();
			}
			catch (Exception ex)
			{
				Log(ex.ToString());
				if (CommonConfig.IS_DEBUG)
				{
					MessageBox.Show(ex.ToString());
				}
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
			try
			{
				if (dbCommand != null)
				{
					dbCommand.Cancel();
				}
			}
			catch (Exception)
			{
			}
			dbCommand = null;
			try
			{
				if (dataReader != null)
				{
					dataReader.Close();
				}
			}
			catch (Exception)
			{
			}
			dataReader = null;
			try
			{
				if (dbConnection != null)
				{
					dbConnection.Close();
				}
			}
			catch (Exception)
			{
			}
			dbConnection = null;
		}

		public SqlCeDataReader ReadFullTable(string tableName)
		{
			string queryString = "SELECT * FROM [" + tableName + "]";
			return ExecuteQuery(queryString);
		}

		public SqlCeDataReader InsertValues(string tableName, string[] values)
		{
			int fieldCount = ReadFullTable(tableName).FieldCount;
			int num = values.Length;
			string text = "INSERT INTO [" + tableName + "] VALUES ('" + values[0] + "'";
			for (int i = 1; i < values.Length; i++)
			{
				text = text + ", '" + values[i] + "'";
			}
			text += " )";
			return ExecuteQuery(text);
		}

		public SqlCeDataReader InsertValues_WithC(string tableName, string colName, string[] values)
		{
			string text = "INSERT INTO [" + tableName + "](" + colName + ") VALUES ('" + values[0] + "'";
			for (int i = 1; i < values.Length; i++)
			{
				text = text + ", '" + values[i] + "'";
			}
			text += " )";
			bool iS_DEBUG = CommonConfig.IS_DEBUG;
			return ExecuteQuery(text);
		}

		public SqlCeDataReader UpdateValues(string tableName, string[] colNames, string[] colValues, string key, string value, string operation)
		{
			int num = colNames.Length;
			int num2 = colValues.Length;
			string text = "UPDATE [" + tableName + "] SET " + colNames[0] + "='" + colValues[0] + "'";
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
			string text = "UPDATE [" + tableName + "] SET " + colNames[0] + "='" + colValues[0] + "'";
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
			string text = "DELETE FROM [" + tableName + "] WHERE " + colNames[0] + operations[0] + "'" + colValues[0] + "'";
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
			string text = "DELETE FROM [" + tableName + "] WHERE " + colNames[0] + operations[0] + "'" + colValues[0] + "'";
			for (int i = 1; i < colValues.Length; i++)
			{
				text = text + " AND " + colNames[i] + operations[i] + "'" + colValues[i] + "'";
			}
			return ExecuteQuery(text);
		}

		public SqlCeDataReader CreateTable(string tableName, string[] colNames, string[] colTypes)
		{
			string text = "CREATE TABLE IF NOT EXISTS [" + tableName + "]( " + colNames[0] + " " + colTypes[0];
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
			text = text + " FROM [" + tableName + "] WHERE " + colNames[0] + " " + operations[0] + " " + colValues[0];
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
