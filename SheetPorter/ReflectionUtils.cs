using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SheetPorter
{
    public static class ReflectionUtils
    {
        public static object CallGeneric(MethodInfo target, Type[] generics, object targetObject = null, object[] parameters = null)
        {
            var genericMethod = target.MakeGenericMethod(generics);
            return genericMethod.Invoke(targetObject, parameters);
        }
    }
}
