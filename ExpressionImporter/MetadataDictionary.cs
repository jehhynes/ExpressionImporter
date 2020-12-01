using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionImporter
{
    public class MetadataDictionary<TDomain> : Dictionary<string, string>
        where TDomain : class
    {
        static readonly EvaluateConstantsExpressionVisitor VISITOR = new EvaluateConstantsExpressionVisitor();
        public string Get<TProp>(Expression<Func<TDomain, TProp>> expression)
        {
            var key = ((LambdaExpression)VISITOR.Visit(expression)).Body.ToString();
            if (this.ContainsKey(key))
            {
                return this[key];
            }
            else
            {
                return null;
            }
        }

        public string Get(IImportProperty prop)
        {
            if (this.ContainsKey(prop.ExpressionKey))
            {
                return this[prop.ExpressionKey];
            }
            else
            {
                return null;
            }
        }
    }
    public interface IHasMetadataRow<TDomain>
        where TDomain : class
    {
        MetadataDictionary<TDomain> Metadata { get; set; }
    }
}
