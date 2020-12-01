using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ExpressionImporter
{
    public interface IImportProperty
    {
        string ExpressionKey { get; }
        Type PropertyType { get; }
        bool Required { get; }
        string FieldName { get; }
        bool ExportOnly { get; }
        string Description { get; }
        bool Update { get; }
    }

    public abstract class ImportProperty<TDomain, TContext> : IImportProperty
        where TDomain : class
    {
        public string ExpressionKey { get; protected set; }
        public LambdaExpression Expression { get; protected set; }
        public Type PropertyType { get; protected set; }

        public bool Required { get; set; }
        public string FieldName { get; set; }
        public bool Update { get; set; }
        public bool ExportOnly { get; set; }
        public Regex ValidateRegex { get; set; }
        public string ExcelNumberFormat { get; set; }
        public string Description { get; set; }

        public abstract object Parse(object value);

        public override string ToString()
        {
            return FieldName;
        }

        protected Lazy<Func<GetValueArgs<TDomain, TContext>, object>> _defaultGetValueFunc;
        public Func<GetValueArgs<TDomain, TContext>, object> GetValueFunc { get; set; }

        public Func<TDomain, bool> RequiredIf { get; set; }
        public Func<TDomain, bool> IgnoreIf { get; set; }

        public abstract void SetValue(TDomain record, bool isNew, object value, TContext context, ImportValueDictionary<TDomain> values);

        public object GetValue(GetValueArgs<TDomain, TContext> e)
        {
            var func = GetValueFunc ?? _defaultGetValueFunc.Value;
            return func.Invoke(e);
        }

        public bool IsIgnoredFor(TDomain domain)
        {
            return IgnoreIf != null && IgnoreIf(domain);
        }
    }

    public partial class ImportProperty<TDomain, TProp, TContext> : ImportProperty<TDomain, TContext>
         where TDomain : class
    {
        protected Lazy<Action<SetValueArgs<TDomain, TProp, TContext>>> _defaultSetValueAction;
        public Action<SetValueArgs<TDomain, TProp, TContext>> SetValueAction { get; set; }
        public Func<TProp, TProp> ParseFunction { get; set; }
        public Func<TProp, TProp> FormatFunction { get; set; }

        public override void SetValue(TDomain record, bool isNew, object value, TContext context, ImportValueDictionary<TDomain> values)
        {
            TProp val = default;
            try
            {
                val = (TProp)value;
            }
            catch (NullReferenceException) when (!Required)
            {
                //Eat exception and return. It's not required so it should stay unchanged, or the default value.
                return;
            }

            var action = SetValueAction ?? _defaultSetValueAction.Value;

            action.Invoke(new SetValueArgs<TDomain, TProp, TContext>() { Record = record, IsNew = isNew, Value = val, Context = context, Values = values });
        }

        public override object Parse(object value)
        {
            if (ParseFunction != null) return ParseFunction((TProp)value);
            else return value;
        }

        static readonly EvaluateConstantsExpressionVisitor VISITOR = new EvaluateConstantsExpressionVisitor();

        public ImportProperty(Expression<Func<TDomain, TProp>> expression, string fieldName = null)
        {
            Expression = expression;
            ExpressionKey = ((LambdaExpression)VISITOR.Visit(expression)).Body.ToString();
            FieldName = fieldName;
            PropertyType = typeof(TProp);

            if (expression.Body is MemberExpression me && me.Member.MemberType == MemberTypes.Property)
            {
                if (fieldName == null)
                    FieldName = me.Member.Name;

                _defaultSetValueAction = new Lazy<Action<SetValueArgs<TDomain, TProp, TContext>>>(() =>
                {
                    var pi = me.Member as PropertyInfo;

                    if (me.Expression == null || me.Expression.NodeType == ExpressionType.Parameter || me.Expression.NodeType == ExpressionType.Convert)
                    {
                        return (e) =>
                        {
                            if (pi.DeclaringType.IsAssignableFrom(e.Record.GetType()))
                            {
                                pi.SetValue(e.Record, e.Value);
                            }
                            else if (e.Value != null)
                            {
                                throw new Exception($"{FieldName} supplied but is non-applicable for type '{e.Record.GetType().Name}'");
                            }
                        };
                    }
                    else
                    {
                        var lambdaParent = System.Linq.Expressions.Expression.Lambda<Func<TDomain, object>>(me.Expression, expression.Parameters);
                        var compiledParent = lambdaParent.Compile();
                        return (e) =>
                        {
                            var parentObj = compiledParent(e.Record);
                            pi.SetValue(parentObj, e.Value);
                        };
                    }
                });
            }

            _defaultGetValueFunc = new Lazy<Func<GetValueArgs<TDomain, TContext>, object>>(() =>
            {
                Lazy<Func<TDomain, TProp>> compiledExpression = new Lazy<Func<TDomain, TProp>>(expression.Compile);
                return (e) =>
                {
                    try
                    {
                        var value = compiledExpression.Value(e.Record);

                        if (FormatFunction == null) return value;
                        else return FormatFunction(value);
                    }
                    catch (NullReferenceException)
                    {
                        return null;
                    }
                    catch (InvalidCastException)
                    {
                        return null;
                    }
                };
            });

        }
    }

    public class SetValueArgs<TDomain, TProp, TContext>
        where TDomain : class
    {
        public TDomain Record { get; set; }
        public bool IsNew { get; set; }
        public TProp Value { get; set; }
        public TContext Context { get; set; }

        public ImportValueDictionary<TDomain> Values { get; set; }

        public TValue GetValue<TValue>(Expression<Func<TDomain, TValue>> expression) => Values.Get<TValue>(expression);
    }

    public class GetValueArgs<TDomain, TContext>
    {
        public TDomain Record { get; set; }
        public TContext Context { get; set; }
    }
}
