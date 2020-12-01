# ExpressionImporter
Framework for importing and exporting data based on Linq expressions. 
This framework allows you to create imports, exports, or an export / re-import workflow using Linq expressions.

## 1. Implement the AbstractImport<TDomain, TId, TContext> base class
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
    
## 2. Create an import class
    public class HelpPageImport : MyAbstractImport<HelpPage>
    {
        public HelpPageImport(MyContext context) : base(context)
        {
            Prop(x => x.Name);
            Prop(x => x.Html);
            Prop(x => x.IsInactive);
        }
    }
    
## 3. Use the import
    using OfficeOpenXml;
    using System.IO;
    
    public void ImportHelpPages(Stream inputStream)
    {
        var import = new HelpPageImport(MyContext);
        var excelPackage = new ExcelPackage(inputStream);
        import.ImportFromExcel(ImportType.Create, excelPackage);
    }


