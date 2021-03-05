﻿using System;
using System.Data.SqlClient;
using System.Xml.Serialization;
using System.IO;
using System.Text;
using System.Xml;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClassLibrarys.Common;
using ApplicationService;

namespace Repository
{
    /// <summary>
    /// データベース接続管理クラス (SqlServer)
    /// </summary>
    public class DBManeger : IDBManeger
    {

        #region 接続情報クラス

        /// <summary>
        /// DB接続情報
        /// </summary>
        public class DBInfomation
        {
            /// <summary>
            /// サーバー名
            /// </summary>
            public string ServerName { get; set; }

            /// <summary>
            /// データベース名
            /// </summary>
            public string DBName { get; set; }

            /// <summary>
            /// ログイン情報
            /// </summary>
            public string Login { get; set; }

            /// <summary>
            /// パスワード
            /// </summary>
            public string Password { get; set; }

            /// <summary>
            /// 接続情報形式で取得する
            /// </summary>
            /// <returns>接続情報を文字列で返す</returns>
            public string CreateConnectString()
                => $@"Data Source = {ServerName};
                      Database={DBName};
                      Persist Security Info=True;
                      User ID={Login};Password={Password}";
        }

        #endregion

        #region メンバー

        /// <summary>
        /// データベース接続情報保存パス
        /// </summary>
        private static readonly string DBInfoPass = "test.xml";

        /// <summary>
        /// Sqlコネクション情報
        /// </summary>
        private static SqlConnection Connection { get; set; }

        /// <summary>
        /// SQLトランザクション
        /// </summary>
        private static SqlTransaction SqlTransaction { get; set; }

        /// <summary>
        /// 接続情報
        /// </summary>
        private static DBInfomation Setting { get; set; }

        /// <summary>
        /// 接続数
        /// </summary>
        protected static int ConnectCount { get; set; }

        //接続先の命名規則
        protected enum DBColumnNameRole
        {
            PascalCase,
            SnakeCase
        }

        /// <summary>
        /// 接続先の命名規則設定
        /// </summary>
        protected DBColumnNameRole ColumnNameRole { get; set; } = DBColumnNameRole.PascalCase;

        /// <summary>
        /// タイムアウト時間
        /// </summary>
        protected readonly int TimeOut = 60000;

        /// <summary>
        /// ユーザーステータス情報
        /// </summary>
        protected IUserState UserState { get; set; }

        #endregion

        #region スタティックイニシャライザ

        /// <summary>
        /// staticイニシャライザ
        /// </summary>
        static DBManeger()
        {
            // 接続情報を読み込む
            Setting = CreateDBInfomation();

            // 接続情報が存在しない場合は接続情報を書き込む
            OutPutDBInfo();
        }
        #endregion

        #region 接続関係

        /// <summary>
        /// DBに接続をする
        /// </summary>
        public void Connect()
        {
            // すでにコネクションが存在すれば接続をしない
            if (Connection is object) return;

            Connection = new SqlConnection();
            Connection.ConnectionString = Setting.CreateConnectString();
            Connection.Open();
        }

        /// <summary>
        /// DB接続情報を作成する
        /// </summary>
        /// <returns> 設定情報を戻す (初期設定情報を戻す) </returns>
        private static DBInfomation CreateDBInfomation()
        {
            try
            {
                using (var streamReader = new StreamReader(DBInfoPass, Encoding.UTF8))
                {
                    var xmlSettings = new XmlReaderSettings();
                    using (var xmlReader = XmlReader.Create(streamReader, xmlSettings))
                    {
                        var xmlDeSerializer = new XmlSerializer(typeof(DBInfomation));
                        return xmlDeSerializer.Deserialize(xmlReader) as DBInfomation;
                    }
                }
            }
            catch (Exception)
            {
                return CreateDBInfo();
            }
        }

        /// <summary>
        /// DB接続情報について出力する
        /// </summary>
        private static void OutPutDBInfo()
        {
            try
            {
                if (File.Exists(DBInfoPass) == false)
                {
                    var xmlSerializer1 = new XmlSerializer(typeof(DBInfomation));
                    using (var streamWriter = new StreamWriter(DBInfoPass, false, Encoding.UTF8))
                    {
                        xmlSerializer1.Serialize(streamWriter, Setting);
                        streamWriter.Flush();
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

#if DEBUG

        /// <summary>
        /// 初回DB接続情報を取得する
        /// </summary>
        /// <returns></returns>
        private static DBInfomation CreateDBInfo()
            => new DBInfomation()
            {
                ServerName = "",
                DBName = "",
                Login = "",
                Password = ""
            };
#else

        /// <summary>
        /// 初回DB接続情報を取得する
        /// </summary>
        /// <returns></returns>
        private static DBInfomation CreateDBInfo()
        => new DBInfomation()
            {
                ServerName = "",
                DBName = "",
                Login = "",
                Password = ""
            };
#endif

        /// <summary>
        /// 接続情報の解放
        /// </summary>
        public void Close()
        {
            // 接続数を1減らす
            ConnectCount--;

            // 他にも接続しているオブジェクトが存在していれば解放しない
            if (ConnectCount != 0) return;

            // コネクションが存在しない場合は実行しない
            if (Connection is null) return;

            // 接続情報を解放する
            Connection.Close();
            Connection.Dispose();
            Connection = null;
        }

        #endregion

        #region SQL実行関係

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public DBManeger()
        {
            Connect();
        }

        /// <summary>
        /// メモリ解放処理
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        /// <summary>
        /// トランザクションを開始する
        /// </summary>
        public void Tran()
            => SqlTransaction = Connection?.BeginTransaction();

        /// <summary>
        /// 処理をコミットする
        /// </summary>
        public void Commit()
            => SqlTransaction?.Commit();

        /// <summary>
        /// ロールバックする
        /// </summary>
        public void RolleBack()
            => SqlTransaction?.Rollback();

        /// <summary>
        /// 戻り値ありクエリ実行
        /// </summary>
        /// <param name="query"> 実行クエリ </param>
        /// <returns> クエリ実行結果 </returns>
        public void ExecuteQuery(DataTable dt,string query)
            => ExecuteQuery(dt,query, new Dictionary<string, object>());

        /// <summary>
        /// 戻り値ありクエリ実行 (コマンドパラメータ辞書型ver)
        /// </summary>
        /// <param name="query"> 実行クエリ </param>
        /// <returns> クエリ実行結果 </returns>
        public void ExecuteQuery(DataTable dt, string query, Dictionary<string, object> param)
            => ExecuteQuery(dt,query, ToSqlParameters(param));

        /// <summary>
        /// 戻り値ありクエリ実行 (コマンドパラメータモデルベース ver)
        /// </summary>
        /// <typeparam name="T"> モデルの型 </typeparam>
        /// <param name="query"> 検索クエリ </param>
        /// <param name="model"> 埋め込むデータ </param>
        /// <returns> クエリ実行結果 </returns>
        public void ExecuteQuery<T>(DataTable dt, string query, T model) where T:class
            => ExecuteQuery(dt,query, GetSqlParameterByPropety<T>(model, typeof(T).GetProperties()));


        /// <summary>
        /// 戻り値ありクエリ実行 (モデルにバインドしたリストで戻す ver)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <returns></returns>
        public List<T> ExecuteQuery<T>(string query)where T:class
            => ToModelsByDataTable<T>(ExecuteQuery(query, null));

        /// <summary>
        /// 戻り値ありクエリ実行 (モデルにバインドしたリストで戻す ver)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <returns></returns>
        public List<T> ExecuteQuery<T>(string query, Dictionary<string, object> param)where T:class
            => ToModelsByDataTable<T>(ExecuteQuery(query, ToSqlParameters(param)));

        /// <summary>
        /// 戻り値ありクエリ実行 (モデルにバインドしたリストで戻す ver)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <returns></returns>
        public List<T> ExecuteQuery<T>(string query, T model) where T:class
            => ToModelsByDataTable<T>(ExecuteQuery(query, GetSqlParameterByPropety<T>(model, typeof(T).GetProperties())));

        /// <summary>
        /// DataTableをモデルリストに変換する
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        private List<T> ToModelsByDataTable<T>(DataTable dt) where T : class
            => BinndingModelByDataTable<T>(GetDataRows(dt),
                                           GetDataTableColumnNames(dt),
                                           GetPropetyNames<T>());

        /// <summary>
        /// IEnumerable<DataRow>をモデルにバインドする
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="drs"></param>
        /// <param name="columnNames"></param>
        /// <param name="propetyNames"></param>
        /// <returns></returns>
        private List<T> BinndingModelByDataTable<T>(IEnumerable<DataRow> drs,
                                                    List<string> columnNames,
                                                    List<string> propetyNames) where T : class
            => drs.Select(m => BindingModelByDataRow<T>(m, columnNames, propetyNames)).ToList();


        /// <summary>
        /// DataTableからDataRowの反復形式を取得する
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private IEnumerable<DataRow> GetDataRows(DataTable dt)
            => dt.Rows.Cast<DataRow>();

        /// <summary>
        /// DataTableからカラム名のリストを取得する
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private List<string> GetDataTableColumnNames(DataTable dt)
            => dt.Columns.Cast<DataColumn>().Select(m => m.ColumnName).ToList();

        /// <summary>
        /// クラスに属するプロパティ名を取得する
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private List<string> GetPropetyNames<T>() where T : class
            => (Activator.CreateInstance(typeof(T))).GetType().GetProperties().Select(m => m.Name).ToList();

        /// <summary>
        /// 戻り値ありクエリ実行
        /// </summary>
        /// <param name="query"> 検索クエリ </param>
        /// <returns> クエリ実行結果 </returns>
        private DataTable ExecuteQuery(string query, IEnumerable<SqlParameter> parameters)
        {
            var result = new DataTable();
            ExecuteQuery(result, query, parameters);
            return result;
        }

        /// <summary>
        /// 戻り値ありクエリ実行
        /// </summary>
        /// <param name="query"> 検索クエリ </param>
        /// <returns> クエリ実行結果 </returns>
        private void ExecuteQuery(DataTable dt,string query, IEnumerable<SqlParameter> parameters)
        {
            //  Sql実行
            using (var command = new SqlCommand(query, Connection))
            {
                // command設定
                command.CommandTimeout = TimeOut;

                // トランザクション設定
                command.Transaction = SqlTransaction;

                // コマンドパラメータ設定
                if (parameters is object) foreach (var param in parameters) command.Parameters.Add(param);


                // SQL実行
                dt.Load(command.ExecuteReader());
            }
        }

        /// <summary>
        /// 取得したDataRowをModelにバインドする
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dr"></param>
        /// <param name="target"></param>
        /// <returns></returns>
        private T BindingModelByDataRow<T>(DataRow dr, List<string> columnNames, List<string> propetyNames) where T : class
        {
            var resultInstance = Activator.CreateInstance(typeof(T)) as T;

            foreach (var columnName in columnNames)
            {
                var propetyName = (ColumnNameRole == DBColumnNameRole.SnakeCase) ? StringProcess.ToCamelBySnake(columnName) : columnName;
                if (propetyNames.Any(m => m == propetyName))
                {
                    resultInstance.GetType().GetProperty(propetyName).SetValue(resultInstance, dr[columnName] is DBNull ? null : dr[columnName]);
                }
            }
            return resultInstance;
        }

        /// <summary>
        /// 戻り値なしクエリ実行
        /// </summary>
        /// <param name="query"実行クエリ></param>
        public void ExecuteNonQuery(string query)
            => ExecuteNonQuery(query, new Dictionary<string, object>());

        /// <summary>
        /// 戻り値なしクエリ実行 (コマンドパラメータあり)
        /// </summary>
        /// <param name="query"> クエリ </param>
        /// <param name="param"> コマンドパラメータ </param>
        public void ExecuteNonQuery(string query, Dictionary<string, object> param)
            => ExecuteNonQuery(query, ToSqlParameters(param));

        /// <summary>
        /// 戻り値なしクエリ実行 (モデルのプロパティ名)
        /// </summary>
        /// <typeparam name="T">モデルの型</typeparam>
        /// <param name="query">実行クエリ</param>
        /// <param name="model">登録対象の情報</param>
        public void ExecuteNonQuery<T>(string query, T model) where T:class
            => ExecuteNonQuery(query, GetSqlParameterByPropety<T>(model, typeof(T).GetProperties()));

        /// <summary>
        /// 戻り値無しクエリ実行 (実行用)
        /// </summary>
        /// <param name="query"> 実行クエリ </param>
        /// <param name="parameters"> コマンドパラメータ </param>
        private void ExecuteNonQuery(string query, IEnumerable<SqlParameter> parameters)
        {
            //  Sql実行
            using (var command = new SqlCommand(query, Connection))
            {
                // command設定
                command.CommandTimeout = TimeOut;

                // トランザクション設定
                command.Transaction = SqlTransaction;

                // パラメータを追加する
                if (parameters is object) foreach (var param in parameters) command.Parameters.Add(param);

                command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// コマンドにパラメータを追加する
        /// </summary>
        /// <param name="command"> 追加対象のコマンド </param>
        /// <param name="param"> 追加するSQLパラメータ </param>
        private IEnumerable<SqlParameter> ToSqlParameters(Dictionary<string, object> param)
            => param.Select(m => new SqlParameter() { ParameterName = m.Key, Value = m.Value });

        /// <summary>
        /// モデルクラスからパラメータクラスを生成する
        /// 使用方法 パラメータは必ずモデルのプロパティと一致させておく必要がある
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="target"></param>
        private IEnumerable<SqlParameter> GetSqlParameterByPropety<T>(T target, IEnumerable<PropertyInfo> propetyNames) where T : class
            => propetyNames.Select(m => new SqlParameter()
            {
                ParameterName = $"@{m.Name}",
                Value = typeof(T).GetProperty(m.Name).GetValue(target)
            });


        #endregion
    }

}
