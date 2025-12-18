using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Heater_Cal_Demo_P4.Data;
using System.Data;
using System.Data.OleDb;

namespace Heater_Cal_Demo_P4.Communication
{
    public sealed class OleDbHelper 
    {
        #region Make Connection String
        public static class ConnStr
        {
            public static string SqlServer(string ip, string port, string dbName, string user, string pw)
            {
                return string.Format(
                   "Provider=SQLOLEDB;Data Source={0},{1};Initial Catalog={2};User ID={3};Password={4};",
                ip, port, dbName, user, pw);
            }

            // 원본 호환(Initial Catalog + database= 옵션 동시 포함)
            public static string SqlServerCompatLegacy(string ip, string port, string dbName, string dbOption, string user, string pw)
            {
                return string.Format(
                    "Provider=SQLOLEDB;Data Source={0},{1};Initial Catalog={2};database={3};User ID={4};Password={5};",
                    ip, port, dbName, dbOption, user, pw);
            }

            public static string AccessAccdb(string path)
            {
                return string.Format(
                    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=False;",
                    path);
            }
        }
        #endregion

        #region Sycn Function
        public static DataSet ExecuteDataSet(string cs, string sql, int timeoutSeconds = 30, params OleDbParameter[] parameters)
        {
            DataSet ds = new DataSet();
            using (var conn = new OleDbConnection(cs))
            using (var cmd = new OleDbCommand(sql, conn))
            using (var adp = new OleDbDataAdapter(cmd))
            {
                cmd.CommandTimeout = timeoutSeconds;
                if (parameters != null)
                {
                    foreach (var p in parameters) cmd.Parameters.Add(p);
                }
                conn.Open();
                adp.Fill(ds);
            }
            return ds;
        }

        public static DataTable ExecuteDataTable(string cs, string sql, int timeoutSeconds = 30, params OleDbParameter[] parameters)
        {
            var ds = ExecuteDataSet(cs, sql, timeoutSeconds, parameters);
            return ds.Tables.Count > 0 ? ds.Tables[0] : new DataTable();
        }

        public static object ExecuteScalar(string cs, string sql, int timeoutSeconds = 30, params OleDbParameter[] parameters)
        {
            using (var conn = new OleDbConnection(cs))
            using (var cmd = new OleDbCommand(sql, conn))
            {
                cmd.CommandTimeout = timeoutSeconds;
                if (parameters != null)
                {
                    foreach (var p in parameters) cmd.Parameters.Add(p);
                }
                conn.Open();
                return cmd.ExecuteScalar();
            }
        }

        public static int ExecuteNonQuery(string cs, string sql, int timeoutSeconds = 30, params OleDbParameter[] parameters)
        {
            using (var conn = new OleDbConnection(cs))
            using (var cmd = new OleDbCommand(sql, conn))
            {
                cmd.CommandTimeout = timeoutSeconds;
                if (parameters != null)
                {
                    foreach (var p in parameters) cmd.Parameters.Add(p);
                }
                conn.Open();
                return cmd.ExecuteNonQuery();
            }
        }

        #endregion
        
        #region Async Function
        public static Task<DataSet> ExecuteDataSetAsync(string cs, string sql, int timeoutSeconds = 30, CancellationToken ct = default(CancellationToken), params OleDbParameter[] parameters)
        {
            return Task.Run<DataSet>(() =>
            {
                DataSet ds = new DataSet();
                using (var conn = new OleDbConnection(cs))
                using (var cmd = new OleDbCommand(sql, conn))
                using (var adp = new OleDbDataAdapter(cmd))
                {
                    cmd.CommandTimeout = timeoutSeconds;
                    if (parameters != null)
                    {
                        foreach (var p in parameters) cmd.Parameters.Add(p);
                    }

                    using (ct.Register(delegate { try { cmd.Cancel(); } catch { } }))
                    {
                        conn.Open();
                        adp.Fill(ds);
                    }
                }
                return ds;
            }, ct);
        }

        public static Task<object> ExecuteScalarAsync(string cs, string sql, int timeoutSeconds = 30, CancellationToken ct = default(CancellationToken), params OleDbParameter[] parameters)
        {
            return Task.Run<object>(() =>
            {
                using (var conn = new OleDbConnection(cs))
                using (var cmd = new OleDbCommand(sql, conn))
                {
                    cmd.CommandTimeout = timeoutSeconds;
                    if (parameters != null)
                    {
                        foreach (var p in parameters) cmd.Parameters.Add(p);
                    }

                    using (ct.Register(delegate { try { cmd.Cancel(); } catch { } }))
                    {
                        conn.Open();
                        return cmd.ExecuteScalar();
                    }
                }
            }, ct);
        }

        public static Task<int> ExecuteNonQueryAsync(string cs, string sql, int timeoutSecdons = 30, CancellationToken ct = default(CancellationToken), params OleDbParameter[] parameters)
        {
            return Task.Run<int>(() =>
            {
                using (var conn = new OleDbConnection(cs))
                using (var cmd = new OleDbCommand(sql, conn))
                {
                    cmd.CommandTimeout = timeoutSecdons;
                    if (parameters != null)
                    {
                        foreach (var p in parameters) cmd.Parameters.Add(p);
                    }

                    using (ct.Register(delegate { try { cmd.Cancel(); } catch { } }))
                    {
                        conn.Open();
                        return cmd.ExecuteNonQuery();
                    }
                }
            }, ct);
        }
        #endregion
    }
}
