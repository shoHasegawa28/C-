using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ClassLibrarys.Common
{
    public class StringProcess
    {
        #region パスカル関係

        /// <summary>
        /// キャメルケースからスネークケースに変換する
        /// 例: TestCase => test_case
        /// </summary>
        /// <returns></returns>
        public static string ToSnakeCaseByCamel(string camelcase)
        {
            // 引数の値が無効値の場合
            if (camelcase.Length <= 0 || camelcase is null) throw new NullReferenceException();

            // ワードごとに区切りリスト型に変換する
            List<string> words = new List<string>();
            for (int i = 1; i <= GetWordNum(camelcase); i++)
            {
                words.Add(SelectIndexUpperWord(camelcase, i));
            }

            return string.Join("_", words).ToLower();
        }

        /// <summary>
        /// 指定した添え字に対応する単語を取得する (パスカル基準)
        /// </summary>
        /// <param name="target"></param>
        /// <param name="upperIndex"></param>
        /// <returns></returns>
        public static string SelectIndexUpperWord(string target, int upperIndex)
        {
            string result = "";
            for (int i = 0, cnt = 0; i < target.Length; i++)
            {
                // 指定した添え字よりもワード数が大きい場合は終了する
                if (cnt > upperIndex) break;

                // 大文字を検知したら文字数カウントを1増やす
                if (Char.IsUpper(target[i])) cnt++;

                // 文字数と指定した添え字が等しい場合は戻り値用の文字列に追加する
                if (upperIndex == cnt) result += target[i];
            }

            return result;
        }

        /// <summary>
        /// 文字列の単語の数を取得する (パスカル基準)
        /// </summary>
        /// <returns></returns>
        public static int GetWordNum(string target)
            => target.Count(m => char.IsUpper(m));

        #endregion

        #region スネークケース
        
        /// <summary>
        /// スネークケースからキャメルケースに変換する
        /// </summary>
        /// <returns></returns>
        public static string ToCamelBySnake(string target)
        {
            if (target is null) return target;

            var array = target.Split('_');
            for (int i = 0;i < array.Length;i++)
            {
                array[i] = IndexOfReplaceChar(0,array[i],char.ToUpper(array[i][0]));
            }

            return string.Join("", array);
        }

        /// <summary>
        /// 指定した位置の文字を置換する
        /// </summary>
        /// <returns></returns>
        public static string IndexOfReplaceChar(int index,string target,char replace)
        {
            if (target.Length < index || target is null) return target;

            var temp = target.ToArray();
            temp[index] = replace;
            return new string(temp);
        }


        #endregion


    }
}
