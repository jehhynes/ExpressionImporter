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
    public partial class ImportProperty<TDomain, TProp, TContext> : IImportProperty
        where TDomain : class
    {
        public class Builder
        {
            readonly ImportProperty<TDomain, TProp, TContext> p;

            public Builder(ImportProperty<TDomain, TProp, TContext> property)
            {
                p = property;
            }

            public Builder Required()
            {
                if (p.IgnoreIf != null)
                    throw new ArgumentException("Cannot supply both Required and IgnoreIf because Required is evaluated before the record is processed while IgnoreIf is evaluated after the record is processed. This would result in fields that should be ignored still being required.");

                p.Required = true;
                return this;
            }

            public Builder RequiredIf(Func<TDomain, bool> requiredIf)
            {
                p.RequiredIf = requiredIf;
                return this;
            }

            public Builder SetValue(Action<SetValueArgs<TDomain, TProp, TContext>> setValueAction)
            {
                p.SetValueAction = setValueAction;
                return this;
            }

            public Builder SetValue(Func<TProp, TProp> parseFunction)
            {
                p.ParseFunction = parseFunction;
                return this;
            }

            public Builder NoUpdate()
            {
                p.Update = false;
                return this;
            }

            public Builder Validate(Regex validateRegex)
            {
                p.ValidateRegex = validateRegex;
                return this;
            }

            public Builder Validate(string validatePattern)
            {
                return Validate(new Regex(validatePattern, RegexOptions.Compiled));
            }

            public Builder IgnoreIf(Func<TDomain, bool> ignoreIf)
            {
                if (p.Required)
                    throw new ArgumentException("Cannot supply both Required and IgnoreIf because Required is evaluated before the record is processed while IgnoreIf is evaluated after the record is processed. This would result in fields that should be ignored still being required.");

                p.IgnoreIf = ignoreIf;
                return this;
            }

            public Builder Format(Func<TProp, TProp> formatFunction)
            {
                p.FormatFunction = formatFunction;
                return this;
            }

            public Builder ExcelNumberFormat(string excelNumberFormat)
            {
                p.ExcelNumberFormat = excelNumberFormat;
                return this;
            }

            public Builder ExportOnly()
            {
                p.ExportOnly = true;
                return this;
            }

            public Builder GetValue(Func<GetValueArgs<TDomain, TContext>, object> getValue)
            {
                p.GetValueFunc = getValue;
                return this;
            }

            public Builder Description(string description)
            {
                p.Description = description;
                return this;
            }
        }
    }
}
