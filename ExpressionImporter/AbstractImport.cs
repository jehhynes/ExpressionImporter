using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;

namespace ExpressionImporter
{
    public abstract class AbstractImport<TDomain, TId, TContext> : IImport
        where TDomain : class
    {
        public AbstractImport(TContext context)
        {
            Context = context;

            _compiledIdExpression = IdExpression.Compile();

            if (CanUpdate)
            {
                _idProperty = CreateIdProperty();
                Properties.Add(_idProperty);
            }
        }

        #region Configuration Settings
        public virtual bool CanUpdate { get => true; }
        public virtual bool ByCellReference { get => false; }
        public virtual int RowStart { get => 1; }
        public virtual int? RowEnd { get => null; }
        #endregion
        
        #region Abstractions
        protected abstract ImportProperty<TDomain, TContext> CreateIdProperty();
        public abstract IQueryable<TDomain> Query();
        protected abstract Expression<Func<TDomain, TId>> IdExpression { get; }
        protected abstract TDomain FindById(TId id);
        #endregion

        #region Instance Variables
        protected readonly TContext Context;

        private ImportProperty<TDomain, TContext> _idProperty;

        private Func<TDomain, TId> _compiledIdExpression;
        private TId GetId(TDomain domain) => _compiledIdExpression(domain);
        #endregion

        #region Import Properties
        public virtual List<ImportProperty<TDomain, TContext>> Properties { get; set; } = new List<ImportProperty<TDomain, TContext>>();
        IEnumerable<IImportProperty> IImport.GetProperties() => Properties;

        public ImportProperty<TDomain, TProp, TContext>.Builder Prop<TProp>(Expression<Func<TDomain, TProp>> expression, string fieldName = null)
        {
            var p = new ImportProperty<TDomain, TProp, TContext>(expression, fieldName);

            Properties.Add(p);

            return new ImportProperty<TDomain, TProp, TContext>.Builder(p);
        }
        #endregion

        #region Default Behavior
        public virtual TDomain CreateRecord(ImportValueDictionary<TDomain> values)
        {
            return Activator.CreateInstance<TDomain>();
        }

        public virtual IEnumerable<DataRow> Sort(DataTable dt) => dt.Rows.Cast<DataRow>();

        //Empty implementations
        public virtual void BeforeImport(IEnumerable<ImportValueDictionary<TDomain>> rows) { }
        public virtual void AfterImport(AfterImportArgs<TDomain> e) { }
        public virtual void ValidateRecord(TDomain obj) { }
        public virtual void BeforeProcessRecord(TDomain domain) { }

        public virtual void ProcessRecord(TDomain domain, ImportValueDictionary<TDomain> values, bool isNew)
        {
            List<Exception> exceptions = new List<Exception>();
            foreach (var prop in Properties.Where(x => !x.IsIgnoredFor(domain)))
            {
                if (prop != _idProperty && (prop.Update || isNew))
                {
                    try
                    {
                        prop.SetValue(domain, isNew, values.Get(prop), Context, values);
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(ex);
                    }
                }
            }

            if (exceptions.Any())
                throw new AggregateException(exceptions);
        }
        #endregion

        #region Import Implementation
        public void ImportFromExcel(ImportType importType, ExcelPackage pkg)
        {
            var ws = pkg.Workbook.Worksheets.First();

            DataTable dt;
            if (ByCellReference)
            {
                dt = PivotExcelWorksheetToDataTableByCellReference(ws);
            }
            else
            {
                dt = ExcelWorksheetToDataTable(ws);
            }
            ImportFromDataTable(importType, dt, true);
        }
        public void ImportFromDataTable(ImportType importType, DataTable dt, bool isFromExcel)
        {
            var import = this;

            Func<DataRow, IImportProperty, string> formatRowNumber = (dr, prop) =>
            {
                if (import.ByCellReference)
                {
                    return prop?.Description == null ? null : (" for '" + prop.Description + "'");
                }
                else
                {
                    int row = import.RowStart + dt.Rows.IndexOf(dr) + 1;
                    if (isFromExcel) row++; //Account for header row
                    if (import is IHasMetadataRow<TDomain>) row++;

                    return $" for {(isFromExcel ? "row" : "record")} {row}";
                }
            };

            List<DataRow> rows = import.Sort(dt).ToList();

            Dictionary<DataRow, ImportValueDictionary<TDomain>> valuesDict = new Dictionary<DataRow, ImportValueDictionary<TDomain>>();

            //Parse the values
            foreach (DataRow dr in rows)
            {
                var values = new ImportValueDictionary<TDomain>();
                foreach (var prop in import.Properties)
                {
                    object val = dr[prop.FieldName];
                    if (val == DBNull.Value) val = null;

                    //Data table stores nullable Enum as int. Have to convert it to the underlying enum type
                    if (Nullable.GetUnderlyingType(prop.PropertyType)?.IsEnum == true && val != null)
                        val = Enum.ToObject(Nullable.GetUnderlyingType(prop.PropertyType), val);

                    values[prop.ExpressionKey] = prop.Parse(val);
                }
                valuesDict[dr] = values;
            }

            import.BeforeImport(valuesDict.Values);

            var exceptions = new List<Exception>();
            var importedRecords = new List<ImportedRecord<TDomain>>();
            
            foreach (DataRow dr in rows)
            {
                //skip rows that are all null
                if (dr.Table.Columns.OfType<DataColumn>().All(c => dr.IsNull(c)))
                    continue;

                try
                {
                    var values = valuesDict[dr];

                    TId id = import.CanUpdate ? values.Get(IdExpression) : default(TId);
                    bool idIsNull = EqualityComparer<TId>.Default.Equals(id, default(TId));

                    //Validate the import type
                    if (idIsNull && importType == ImportType.Update)
                        exceptions.Add(new Exception($"Valiation failed{formatRowNumber(dr, null)}: Attempting to update a non-existing record."));
                    else if (!idIsNull && importType == ImportType.Create)
                        exceptions.Add(new Exception($"Valiation failed{formatRowNumber(dr, null)}: Attempting to create an existing record."));

                    //Don't continue to validate this record if there are errors at this point.
                    if (exceptions.Any())
                        continue;

                    //Get the domain object
                    TDomain domain = null;
                    if (!idIsNull) domain = FindById(id);
                    else domain = import.CreateRecord(values);

                    bool isNew = EqualityComparer<TId>.Default.Equals(GetId(domain), default(TId));

                    //Validate required properties
                    foreach (var prop in import.Properties.Where(x => !x.IsIgnoredFor(domain)))
                    {
                        if (prop.Update || isNew) //Do not validate properties if they will not be updated and this is an update import
                        {
                            if (prop.Required && values.Get(prop) == null)
                            {
                                exceptions.Add(new Exception($"Validation failed{formatRowNumber(dr, prop)}: {prop.FieldName} is missing."));
                            }
                        }
                    }

                    //Validate Regex
                    foreach (var prop in import.Properties.Where(x => x.ValidateRegex != null && !x.IsIgnoredFor(domain)))
                    {
                        string str = values.Get(prop)?.ToString();
                        if (str != null)
                        {
                            if (!prop.ValidateRegex.IsMatch(str))
                                exceptions.Add(new Exception($"Validation failed{formatRowNumber(dr, prop)}: {prop.FieldName} must match /{prop.ValidateRegex}/"));
                        }
                    }

                    //Don't process this record if there are validation errors at this point.
                    if (exceptions.Any())
                        continue;

                    BeforeProcessRecord(domain);
                 
                    //Process the row
                    import.ProcessRecord(domain, values, isNew);

                    //Validate the resulting record
                    try
                    {
                        //Validate conditionally required properties
                        foreach (var prop in import.Properties.Where(x => x.RequiredIf != null && !x.IsIgnoredFor(domain)))
                        {
                            if (prop.RequiredIf(domain) && values.Get(prop) == null)
                            {
                                exceptions.Add(new Exception($"Validation failed{formatRowNumber(dr, prop)}: {prop.FieldName} is missing."));
                            }
                        }

                        import.ValidateRecord(domain);
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(new Exception($"Validation failed{formatRowNumber(dr, null)}: {ex.Message}", ex));
                    }

                    importedRecords.Add(new ImportedRecord<TDomain>() { Record = domain, Values = values, IsNew = isNew });
                }
                catch (Exception ex)
                {
                    string messages = ex.Message;
                    if (ex is AggregateException ae)
                        messages = string.Join("; ", ae.InnerExceptions.Select(x => x.Message));
                    exceptions.Add(new Exception($"Import failed{formatRowNumber(dr, null)}: {messages}", ex));
                }
            }

            if (exceptions.Any())
                throw new AggregateException(exceptions);

            import.AfterImport(new AfterImportArgs<TDomain>(importedRecords, importType));
        }
        #endregion

        #region Excel Helpers
        //https://riptutorial.com/epplus/example/27956/create-a-datatable-from-excel-file
        public DataTable ExcelWorksheetToDataTable(ExcelWorksheet worksheet)
        {
            var propTypes = Properties.ToDictionary(x => x.FieldName, x => x.PropertyType, StringComparer.OrdinalIgnoreCase);

            DataTable dt = new DataTable();

            //check if the worksheet is completely empty
            if (worksheet.Dimension == null)
            {
                return dt;
            }

            //create a list to hold the column names
            List<string> columnNames = new List<string>();

            //needed to keep track of empty column headers
            int currentColumn = 1;

            int headerRow = RowStart;

            //loop all columns in the sheet and add them to the datatable
            foreach (var cell in worksheet.Cells[headerRow, 1, headerRow, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();

                //check if the previous header was empty and add it if it was
                if (cell.Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                    currentColumn++;
                }

                //add the column name to the list to count the duplicates
                columnNames.Add(columnName);

                //count the duplicate column names and make them unique to avoid the exception
                //A column named 'Name' already belongs to this DataTable
                int occurrences = columnNames.Count(x => x.Equals(columnName));
                if (occurrences > 1)
                {
                    columnName = columnName + "_" + occurrences;
                }

                //add the column to the datatable
                if (propTypes.ContainsKey(columnName))
                {
                    //data set doesn't like nullables
                    var propType = Nullable.GetUnderlyingType(propTypes[columnName]) ?? propTypes[columnName];

                    dt.Columns.Add(columnName, propType);
                }
                else
                    dt.Columns.Add(columnName);

                currentColumn++;
            }

            var exceptions = new List<Exception>();
            foreach (var prop in Properties)
            {
                if (!dt.Columns.Contains(prop.FieldName))
                {
                    exceptions.Add(new Exception($"Column \"{prop.FieldName}\" is missing."));
                }
            }

            if (exceptions.Any())
                throw new AggregateException(exceptions);

            int firstDataRow = headerRow + 1;
            bool hasMetadataRow = this is IHasMetadataRow<TDomain>;
            if (hasMetadataRow) firstDataRow++;
            int? lastRow = Math.Min(RowEnd ?? worksheet.Dimension.End.Row, worksheet.Dimension.End.Row);

            //start adding the contents of the excel file to the datatable
            for (int i = firstDataRow; i <= lastRow; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();

                //loop all cells in the row
                foreach (var cell in row)
                {
                    if (cell.Start.Column <= dt.Columns.Count)
                    {
                        Type type = null;
                        try
                        {
                            type = dt.Columns[cell.Start.Column - 1].DataType;
                            newRow[cell.Start.Column - 1] = ParseFromExcel(cell.Value, type);
                        }
                        catch (Exception ex)
                        {
                            exceptions.Add(new Exception($"Error parsing {cell.Address} as {type?.Name ?? "?"}", ex));
                        }
                    }
                }

                dt.Rows.Add(newRow);
            }

            if (hasMetadataRow)
            {
                var metaRow = worksheet.Cells[2, 1, 2, worksheet.Dimension.End.Column];

                //pull out the metadata by field name (caption)
                var metasByFieldName = new Dictionary<string, string>();
                foreach (var cell in metaRow)
                {
                    string fieldName;
                    try
                    {
                        fieldName = dt.Columns[cell.Start.Column - 1].Caption;
                        metasByFieldName.Add(fieldName, cell.Value?.ToString());
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(new Exception($"Error parsing metadata in {cell.Address}", ex));
                    }
                }

                //transform to MetadataDictionary
                {
                    var metadata = new MetadataDictionary<TDomain>();

                    foreach (var prop in Properties.Where(x => metasByFieldName.ContainsKey(x.FieldName)))
                        metadata[prop.ExpressionKey] = metasByFieldName[prop.FieldName];

                    (this as IHasMetadataRow<TDomain>).Metadata = metadata;
                }
            }

            if (exceptions.Any())
                throw new AggregateException(exceptions);

            return dt;
        }

        public DataTable PivotExcelWorksheetToDataTableByCellReference(ExcelWorksheet worksheet)
        {
            DataTable dt = new DataTable();

            foreach (var prop in Properties)
            {
                var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                dt.Columns.Add(prop.FieldName, propType);
            }

            var exceptions = new List<Exception>();

            var row = dt.NewRow();
            foreach (var prop in Properties)
            {
                var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                try
                {
                    var cell = worksheet.Cells[prop.FieldName];

                    object value = ParseFromExcel(cell.Value, propType);
                    row[prop.FieldName] = value;
                }
                catch (Exception ex)
                {
                    exceptions.Add(new Exception($"Error parsing {prop.FieldName} as {propType.Name}", ex));
                }
            }
            dt.Rows.Add(row);

            if (exceptions.Any())
                throw new AggregateException(exceptions);

            return dt;
        }
     
        public virtual object ParseFromExcel(object src, Type type)
        {
            if (type == null) return null;

            {
                if (src == null || src == DBNull.Value || (src is string str && string.IsNullOrWhiteSpace(str)))
                    return DBNull.Value;
            }

            if (type == typeof(string))
            {
                return src.ToString();
            }
            else if (type.IsEnum)
            {
                return Enum.Parse(type, src.ToString());
            }
            else if (type == typeof(DateTime) || type == typeof(DateTime?))
            {
                DateTime? dateTime = (src is DateTime || src is DateTime?) ? (DateTime?)src
                    : DateTime.FromOADate((double)src);
                return dateTime;
            }
            else if (type == typeof(int) || type == typeof(int?))
            {
                return Convert.ToInt32(src);
            }
            else if (type == typeof(long) || type == typeof(long?))
            {
                return Convert.ToInt64(src);
            }
            else if (type == typeof(decimal) || type == typeof(decimal?))
            {
                return Convert.ToDecimal(src);
            }
            else if (type == typeof(float) || type == typeof(float?))
            {
                return Convert.ToSingle(src);
            }
            else if (type == typeof(double) || type == typeof(double?))
            {
                return Convert.ToDouble(src);
            }
            else if (type == typeof(bool) || type == typeof(bool?))
            {
                if (src is string str)
                {
                    if (str == "Yes") return true;
                    else if (str == "No") return false;
                }
                else
                    return Convert.ToBoolean(src);
            }

            return src;
        }
        #endregion

        #region Convenience Methods
        protected TProp GetFromLookup<TProp>(ILookup<string, TProp> lookup, string val, string errorPrefix = null)
        {
            if (val != null)
            {
                var matches = lookup[val];
                if (matches.Count() > 1)
                    throw new Exception($"{(errorPrefix == null ? null : (errorPrefix + ": "))}Multiple {typeof(TProp).Name} records found with identifier '{val}'.");
                if (matches.Count() == 0)
                    throw new Exception($"{(errorPrefix == null ? null : (errorPrefix + ": "))}Could not find {typeof(TProp).Name} with identifier '{val}'");

                return matches.Single();
            }
            return default;
        }

        protected TProp GetFromDictionary<TProp>(IDictionary<string, TProp> dict, string val, string errorPrefix = null)
        {
            if (val != null)
            {
                dict.TryGetValue(val, out var record);

                if (record == null)
                    throw new Exception($"{(errorPrefix == null ? null : (errorPrefix + ": "))}Could not find {typeof(TProp).Name} with identifier '{val}'");
                return record;
            }
            return default;
        }
        #endregion
    }

    public class AfterImportArgs<TDomain>
        where TDomain : class
    {
        public AfterImportArgs(IEnumerable<ImportedRecord<TDomain>> records, ImportType importType)
        {
            Records = records;
            ImportType = importType;
        }

        public IEnumerable<ImportedRecord<TDomain>> Records { get; protected set; }
        public ImportType ImportType { get; protected set; }
    }

    public class ImportedRecord<TDomain>
        where TDomain : class
    {
        public TDomain Record { get; set; }
        public ImportValueDictionary<TDomain> Values { get; set; }
        public bool IsNew { get; set; }
    }
}