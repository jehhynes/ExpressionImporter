using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionImporter
{
    /// <summary>Walks your expression and eagerly evaluates property/field members and substitutes them with constants.
    /// You must be sure this is semantically correct, by ensuring those fields (e.g. references to captured variables in your closure)
    /// will never change, but it allows the expression to be compiled more efficiently by turning constant numbers into true constants, 
    /// which the compiler can fold.</summary>
    /// https://stackoverflow.com/questions/3991621/can-i-capture-a-local-variable-into-a-linq-expression-as-a-constant-rather-than
    public class EvaluateConstantsExpressionVisitor : ExpressionVisitor
    {
        protected override Expression VisitMember(MemberExpression m)
        {
            Expression exp = base.Visit(m.Expression);

            if (exp == null || exp is ConstantExpression) // null=static member
            {
                object @object = exp == null ? null : ((ConstantExpression)exp).Value;
                object value = null; Type type = null;
                if (m.Member is FieldInfo fi)
                {
                    value = fi.GetValue(@object);
                    type = fi.FieldType;
                }
                else if (m.Member is PropertyInfo pi)
                {
                    if (pi.GetIndexParameters().Length != 0)
                        throw new ArgumentException("cannot eliminate closure references to indexed properties");
                    value = pi.GetValue(@object, null);
                    type = pi.PropertyType;
                }

                if (type.IsValueType || Nullable.GetUnderlyingType(type)?.IsValueType == true || type == typeof(string)) //JH: Only visit if the constant type is value type or string
                    return Expression.Constant(value, type);
            }

            // otherwise just pass it through
            return Expression.MakeMemberAccess(exp, m.Member);
        }
    }
}
