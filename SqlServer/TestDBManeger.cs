using ApplicationService;
using System;
using System.Collections.Generic;
using System.Data;
using ApplicationService.Model;


namespace TestApplicationFun
{
    public class TestDBManeger
    {
        private IUnitOfWark unitOfWark;
        public TestDBManeger(IUnitOfWark _unitOfWark)
        {
            unitOfWark =_unitOfWark;
        }

        public void TestDBManger()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                // DB接続
                try
                {
                    db.Tran();

                    TestExcuteQuery4();

                    db.Commit();

                }
                catch (Exception)
                {
                    db.RolleBack();
                    throw;
                }
            }
        }

        /// <summary>
        ///　取得クエリテスト (基本)
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var result = new DataTable();
                db.ExecuteQuery(result, "select * from Test1Table");
            }
        }

        /// <summary>
        /// 取得クエリテスト (コマンドパラメータあり (辞書型))
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery2()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var param = new Dictionary<string, object>();
                param.Add("@Id", 2);

                var result = new DataTable();
                db.ExecuteQuery(result, "select * from Test1Table where Id = @Id", param);
            }

        }

        /// <summary>
        /// 取得クエリテスト (コマンドパラメータあり (モデル))
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery3()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var model = new TestModel() { Id = 1, Name = "test" };

                var result = new DataTable();
                db.ExecuteQuery(result, "select * from Test1Table where Id = @Id", model);
            }

        }

        /// <summary>
        /// 取得クエリテスト (コマンドパラメータあり)
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery4()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var result = db.ExecuteQuery<TestModel>("select * from Test1Table");
            }

        }

        /// <summary>
        /// 取得クエリテスト (コマンドパラメータあり)
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery5()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var param = new Dictionary<string, object>();
                param.Add("@Id", 1);

                var result = db.ExecuteQuery<TestModel>("select * from Test1Table where id = @Id", param);
            }

        }

        /// <summary>
        /// 取得クエリテスト (コマンドパラメータあり)
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery6()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var model = new TestModel() { Id = 1, Name = "test" };

                var result = db.ExecuteQuery<TestModel>("select * from Test1Table where id = @Id", model);
            }

        }


        /// <summary>
        /// 取得クエリテスト (DBスネークケースでモデルがパスカルの場合)
        /// </summary>
        /// <returns></returns>
        public void TestExcuteQuery7()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var result = db.ExecuteQuery<TestSnakeCaseModel>("select * from Test2SnakeCase");
            }

        }
        /// <summary>
        /// データ追加クエリ (コマンドパラメータ無し)
        /// </summary>
        public void TestNonExcuteQuery()
        {

            using (IDBManeger db = unitOfWark.DBManeger)
            {
                db.ExecuteNonQuery("Insert into Test1Table values(3,'test')");
            }
        }

        /// <summary>
        /// データ追加クエリ (コマンドパラメータあり)
        /// </summary>
        public void TestNonExcuteQuery2()
        {

            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var param = new Dictionary<string, object>();
                param.Add("@Id", 4);
                db.ExecuteNonQuery("Insert into Test1Table values(@Id,'test')", param);
            }
        }


        /// <summary>
        /// データ追加クエリ (コマンドパラメータあり)
        /// </summary>
        public void TestNonExcuteQuery3()
        {

            using (IDBManeger db = unitOfWark.DBManeger)
            {
                var model = new TestModel() { Id = 1, Name = "test" };
                db.ExecuteNonQuery("Insert into Test1Table values(@Id,@Name)", model);
            }
        }

        /// <summary>
        /// ロールバックについてのテスト
        /// 本来はTesMonExcuteQueryにより3が追加されているがエラーの発生により
        /// ロールバックされて3の追加情報がなくなる
        /// </summary>
        public void TestRollBack()
        {
            using (IDBManeger db = unitOfWark.DBManeger)
            {
                // DB接続
                try
                {
                    db.Tran();

                    TestExcuteQuery();
                    TestExcuteQuery2();

                    TestNonExcuteQuery();

                    // ロールバックテスト用
                    if (true) throw new Exception();

                    TestNonExcuteQuery2();


                    db.Commit();

                }
                catch (Exception)
                {
                    db.RolleBack();
                    throw;
                }
            }
        }
    }
}
