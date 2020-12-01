# ExpressionImporter
Framework for importing and exporting data based on Linq expressions. 
This framework allows you to create imports, exports, or an export / re-import workflow using Linq expressions.

### 1. Implement the AbstractImport<TDomain, TId, TContext> base class
ExpressionImporter is compatible with any ORM or database layer you may happen to be using, including Dapper, NHibernate, or EntityFramework. Simply implement the AbstractImport base class class.

TDomain is the *Type* of object to be imported

TId is the **primary key** *Type* used to identify the object

TContext is a class or interface used to convey any information necessary for processing the import, including a DB Context object or any user context information.

    public abstract class MyAbstractImport<TDomain> : AbstractImport<TDomain, int, MyContext>
        where TDomain : IPersistentObject //Replace IPersistentObject with your object's base class or interface
    {
        public MyAbstractImport(MyContext context) : base(context)
        {
        }

        protected override Expression<Func<TDomain, int>> IdExpression => po => po.Id;

        protected override ImportProperty<TDomain, MyContext> CreateIdProperty()
        {
            return new ImportProperty<TDomain, int?, MyContext>(x => x.Id, "Id") { Update = false };
        }

        protected override TDomain FindById(int id)
        {
            return Context.Session.Get<TDomain>(id);
        }

        public override IQueryable<TDomain> Query()
        {
            return Context.Session.Query<TDomain>();
        }

        public override void AfterImport(AfterImportArgs<TDomain> e)
        {
            //In this example, MyContext is an NHibernate session used to save the new records
            foreach (var record in e.Records.Where(x => x.IsNew).Select(x => x.Record))
                Context.Save(record);
            Context.Flush();
        }
    }
    
### 2. Create an import class
    public class HelpPageImport : MyAbstractImport<HelpPage>
    {
        public HelpPageImport(MyContext context) : base(context)
        {
            Prop(x => x.Name);
            Prop(x => x.Html);
            Prop(x => x.IsInactive);
        }
    }
    
### 3. Use the import
    using OfficeOpenXml;
    using System.IO;
    
    public void ImportHelpPages(Stream inputStream)
    {
        var import = new HelpPageImport(MyContext);
        var excelPackage = new ExcelPackage(inputStream);
        import.ImportFromExcel(ImportType.Create, excelPackage);
    }

## Property Options
Use fluent syntax to set options for a property. The following make the Name property be required:

    Prop(x => x.Name).Required()

The following options are supported:
 - **Required**() - Marks a property as required
 - **RequiredIf**(Func<TDomain, bool> requiredIf) - Conditinally require the property
 - **SetValue**(Action<SetValueArgs<TDomain, TProp, TContext>> setValueAction) - Use this if you need specialized logic for setting the value during an import
 - **Parse**(Func<TProp, TProp> parseFunction) - Use this to transform the input value
 - **NoUpdate**() - prevents modifying the value for existing records. The value is only set for new records
 - **Validate**(Regex validateRegex) - Validate a string value using the provided Regex
 - **Validate**(string validatePattern) - Validate a string value using a Regex created from the provided pattern
 - **IgnoreIf**(Func<TDomain, bool> ignoreIf) - Conditionally ignore/skip setting this property
 
## Customizing the Import

The following virtual methods can be overridden by your import to customize the behavior:
- **BeforeImport**(IEnumerable<ImportValueDictionary<TDomain>> rows) - called before the first record is processed.
- **AfterImport**(AfterImportArgs<TDomain> e) - called after the last record is processed.
- **ValidateRecord**(TDomain obj) - used to validate a record after it has been processed. Throw an exception in this method if validation fails.
- **BeforeProcessRecord**(TDomain domain) - called before each record is processed.
- **ProcessRecord**(TDomain domain, ImportValueDictionary<TDomain> values, bool isNew) - override to add customized processing logic
- **CreateRecord**(ImportValueDictionary<TDomain> values) - The default implementation calls Activator.CreateInstance<TDomain>(). Override to provide customized logic to create new objects.
- **Sort**(DataTable dt) - Override if your import requires the incoming data to be sorted before processing.

## Validation

If validation fails, an exception is thrown. You should catch exceptions and handle them appropriately, for instance displaying a message to the end user. If multiple errors are returned, they will be in an AggregateException. The following example (using MVC) writes any validation errors to ModelState, to be returned to the end user. It also rolls back the DB Transaction when an exception is caught.

    public abstract class MyAbstractImport<TDomain> : AbstractImport<TDomain, int, MyContext>
        where TDomain : IPersistentObject 
    {
        //...other implementation details omitted
    
        public void RunImport(Controller controller, ExcelPackage package)
        {
            try
            {
                ImportFromExcel(ImportType.Create, package);

                controller.ViewBag.Success = true;
            }
            catch (Exception ex)
            {
                var exceptions = Flatten(ex);

                Context.Transaction.Rollback();

                controller.ModelState.AddModelError("", $"Your import could not be completed:");

                foreach (var exception in exceptions)
                {
                    controller.ModelState.AddModelError("", exception.Message);
                }
            }
        }

        static IList<Exception> Flatten(Exception ex)
        {
            var exceptions = new List<Exception>();
            if (ex is AggregateException ae)
                exceptions.AddRange(ae.InnerExceptions);
            else
                exceptions.Add(ex);
            return exceptions;
        }
    }

