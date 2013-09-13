using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LinqToExcel.Domain;
using Remotion.Collections;

namespace LinqToExcel.Query
{
    internal class ExcelQueryConstructorArgs
    {
        internal string FileName { get; set; }
        internal DatabaseEngine DatabaseEngine { get; set; }
        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal Dictionary<TransformKey, Func<string, object>> Transformations { get; set; }
        internal Dictionary<Type, Func<string, object>> TypeTransformations { get; set; }
        internal Dictionary<TransformKey, Tuple<string, Func<object, IQueryable, IQueryable>>> ForeignKeyTransformations { get; set; }
        internal Dictionary<TransformKey, IQueryable> SheetDataIncludes { get; set; }
        internal StrictMappingType? StrictMapping { get; set; }
    }
}
