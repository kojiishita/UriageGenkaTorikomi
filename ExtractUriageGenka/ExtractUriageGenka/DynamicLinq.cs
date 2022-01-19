using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExtractUriageGenka
{
    public static class EnumerableExtensions
    {
        /// <summary>
        /// 動的Where句作成(複数)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="expressions">カラム名, 値, 型, "And"・"Or"</param>
        /// <returns></returns>
        public static IEnumerable<T> DynamicWhere<T>(this IEnumerable<T> source, List<Tuple<string, string, object, Type>> expressions)
        {
            var queryableSource = source.AsQueryable();

            // 条件作成
            var lambda = GetPredicate<T>(expressions);

            return queryableSource.Where(lambda);
        }

        /// <summary>
        /// 動的Where句条件作成
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expressions"></param>
        /// <returns></returns>
        private static Func<T, bool> GetPredicate<T>(List<Tuple<string, string, object, Type>> expressions)
        {
            // パラメータの定義する:x
            ParameterExpression param = Expression.Parameter(typeof(T), "x");

            // 全体のbody
            BinaryExpression body = null;

            int index = 0;
            foreach (var exp in expressions)
            {
                // x => x.Id == id
                // Idのbodyの左を定義する：x.Id
                MemberExpression left = Expression.Property(param, exp.Item1);

                // Idのbodyの右を定義する:id
                ConstantExpression right = Expression.Constant(exp.Item3, exp.Item4);

                // Idのbodyを定義する：x.Id == id
                BinaryExpression bodyDetails = exp.Item2 switch
                {
                    "=" => Expression.Equal(left, right),
                    "!=" => Expression.NotEqual(left, right),
                    ">" => Expression.GreaterThan(left, right),
                    "<" => Expression.LessThan(left, right),
                    ">=" => Expression.GreaterThanOrEqual(left, right),
                    "<=" => Expression.LessThanOrEqual(left, right),
                    _ => Expression.Equal(left, right)
                };

                //BinaryExpression bodyDetails = Expression.Equal(left, right);

                // Andでつなぐ
                body = (index == 0) ? bodyDetails : Expression.And(body, bodyDetails);

                index++;
            }

            // 式ツリーを(x => x.Id == id || x => x.Name == name)を組み立て、
            // 実行コードにコンパイルする
            return Expression.Lambda<Func<T, bool>>(body, param).Compile();
        }
    }

    /// <summary>
    /// 条件式
    /// </summary>
    public enum LinqExpression : int
    {
        And = 1,
        Or = 2,
    }
}