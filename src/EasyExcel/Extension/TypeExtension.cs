using System;

namespace EasyExcel.Extension
{
    public static class TypeExtension
    {
        public static bool IsNumbericType(this Type type)
        {
            return type == typeof(int)
            || type == typeof(double)
            || type == typeof(long)
            || type == typeof(short)
            || type == typeof(float)
            || type == typeof(Int16)
            || type == typeof(Int32)
            || type == typeof(Int64)
            || type == typeof(uint)
            || type == typeof(UInt16)
            || type == typeof(UInt32)
            || type == typeof(UInt64)
            || type == typeof(sbyte)
            || type == typeof(Single)
            || type == typeof(decimal);
        }

        public static bool IsNullable(this Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
        }
    }
}