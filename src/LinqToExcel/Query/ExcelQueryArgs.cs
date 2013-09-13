using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LinqToExcel.Domain;
using Remotion.Collections;

namespace LinqToExcel.Query
{
    internal class ExcelQueryArgs
    {
        internal string FileName { get; set; }
        internal DatabaseEngine DatabaseEngine { get; set; }
        internal string WorksheetName { get; set; }
        internal int? WorksheetIndex { get; set; }
        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal Dictionary<TransformKey, Func<string, object>> Transformations { get; private set; }
        internal Dictionary<Type, Func<string, object>> TypeTransformations { get; private set; }
        internal Dictionary<TransformKey, Tuple<string, Func<object, IQueryable, IQueryable>>> ForeignKeyTransformations { get; set; }
        internal Dictionary<TransformKey, IQueryable> SheetDataIncludes { get; set; }
        internal string StartRange { get; set; }
        internal string EndRange { get; set; }
        internal bool NoHeader { get; set; }
        internal StrictMappingType? StrictMapping { get; set; }

        internal ExcelQueryArgs()
            : this(new ExcelQueryConstructorArgs() { DatabaseEngine = ExcelUtilities.DefaultDatabaseEngine() })
        { }

        internal ExcelQueryArgs(ExcelQueryConstructorArgs args)
        {
            FileName = args.FileName;
            DatabaseEngine = args.DatabaseEngine;
            ColumnMappings = args.ColumnMappings ?? new Dictionary<string, string>();
            Transformations = args.Transformations ?? new Dictionary<TransformKey, Func<string, object>>();
            TypeTransformations = args.TypeTransformations ?? new Dictionary<Type, Func<string, object>>();
            ForeignKeyTransformations = args.ForeignKeyTransformations ?? new Dictionary<TransformKey, Tuple<string, Func<object, IQueryable, IQueryable>>>();
            StrictMapping = args.StrictMapping ?? StrictMappingType.None;
            SheetDataIncludes = args.SheetDataIncludes ?? new Dictionary<TransformKey, IQueryable>();
        }

        public override string ToString()
        {
            var columnMappingsString = new StringBuilder();
            foreach (var kvp in ColumnMappings)
                columnMappingsString.AppendFormat("[{0} = '{1}'] ", kvp.Key, kvp.Value);
            var transformationsString = string.Join(", ", Transformations.Keys.Select(tk => tk.Item2).ToArray());
            var typeTransformationsString = string.Join(", ", TypeTransformations.Keys.Select(t => t.Name).ToArray());

            return string.Format("FileName: '{0}'; WorksheetName: '{1}'; WorksheetIndex: {2}; StartRange: {3}; EndRange: {4}; NoHeader: {5}; ColumnMappings: {6}; Transformations: {7}, Transformations: {8}, StrictMapping: {9}",
                FileName, WorksheetName, WorksheetIndex, StartRange, EndRange, NoHeader, columnMappingsString, transformationsString, typeTransformationsString, StrictMapping);
        }

    }
}
