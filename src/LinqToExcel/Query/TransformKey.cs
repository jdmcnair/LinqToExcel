using Remotion.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Query
{
    internal class TransformKey : Tuple<Type, string>
    {
        protected TransformKey(Type type, string propertyName)
            : base(type, propertyName) { }
        public static TransformKey Create(Type type, string propertyName)
        {
            return new TransformKey(type, propertyName);
        }
    }
}
