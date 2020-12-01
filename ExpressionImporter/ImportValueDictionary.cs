using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionImporter
{
    public class ImportValueDictionary<TDomain> : Dictionary<string, object>
        where TDomain : class
    {
        static readonly EvaluateConstantsExpressionVisitor VISITOR = new EvaluateConstantsExpressionVisitor();
        public TProp Get<TProp>(Expression<Func<TDomain, TProp>> expression)
        {
            var key = ((LambdaExpression)VISITOR.Visit(expression)).Body.ToString();
            if (this.ContainsKey(key))
            {
                return (TProp)this[key];
            }
            else
            {
                throw new InvalidOperationException($"Could not find expression {key}");
            }
        }

        public object Get(IImportProperty prop)
        {
            if (this.ContainsKey(prop.ExpressionKey))
            {
                return this[prop.ExpressionKey];
            }
            else
            {
                throw new InvalidOperationException($"Could not find expression {prop.ExpressionKey}");
            }
        }
    }
}
