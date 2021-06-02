using System.Data;
using System.Collections.Generic;
using System;

namespace DBManeger
{
    public interface IDBManeger : IDisposable
    {
        /// <summary>
        /// 接続
        /// </summary>
        public void Connect();

        /// <summary>
        /// 閉じる
        /// </summary>
        public void Close();

        /// <summary>
        /// トランザクション
        /// </summary>
        public void Tran();

        /// <summary>
        /// コミット
        /// </summary>
        public void Commit();

        /// <summary>
        /// ロールバック
        /// </summary>
        public void RolleBack();

        /// <summary>
        /// 戻り値なしSQL実行
        /// </summary>
        public void ExecuteNonQuery(string query);

        /// <summary>
        /// 戻り値なしSQL実行
        /// </summary>
        public void ExecuteNonQuery(string query, Dictionary<string, object> param);

        /// <summary>
        /// 戻り値なしSQL
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <param name="model"></param>
        public void ExecuteNonQuery<T>(string query, T model) where T : class;

        /// <summary>
        /// 戻り値ありSQL実行
        /// </summary>
        public void ExecuteQuery(DataTable dt, string query);

        /// <summary>
        /// 戻り値ありSQL実行
        /// </summary>
        public void ExecuteQuery(DataTable dt, string query, Dictionary<string, object> param);

        /// <summary>
        /// 戻り値ありSQL実行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <param name="model"></param>
        public void ExecuteQuery<T>(DataTable dt, string query, T model) where T : class;

        /// <summary>
        /// 戻り値ありSQL実行
        /// </summary>
        public List<T> ExecuteQuery<T>(string query) where T : class;

        /// <summary>
        /// 戻り値ありSQL実行
        /// </summary>
        public List<T> ExecuteQuery<T>(string query, Dictionary<string, object> param) where T : class;

        /// <summary>
        /// 戻り値ありSQL実行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <param name="model"></param>
        public List<T> ExecuteQuery<T>(string query, T model) where T : class;
    }
}
