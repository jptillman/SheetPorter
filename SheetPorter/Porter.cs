using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SheetPorter.Attributes;

namespace SheetPorter
{
    public static class Porter
    {
        public static IEnumerable<T> Port<T>(Stream source) where T : new()
        {
            var book = new HSSFWorkbook(source);
            var sheet = GetSheet(book, typeof(T));

            var tableAtt = typeof(T).GetCustomAttributes(typeof(TableAttribute), true).Cast<TableAttribute>().FirstOrDefault();
            var propSetAtt = typeof(T).GetCustomAttributes(typeof(PropertySetAttribute), true).Cast<PropertySetAttribute>().FirstOrDefault();

            if (tableAtt == null && propSetAtt == null)
            {
               throw new Exception($"The type {typeof(T).Name} has neither a Table attribute nor a PropertySet attribute.  Cannot port."); 
            }

            return tableAtt != null ? WorksheetToModels<T>(sheet, tableAtt.HeaderMatchPattern) : new [] {WorksheetToModel<T>(sheet)};
        }

        public static IEnumerable<T> Port<T>(Stream source, T model) where T : new()
        {
            return Port<T>(source);
        }

        private static ISheet GetSheet(IWorkbook book, Type targetType)
        {
            var spec =
                targetType.GetCustomAttributes(typeof(TableAttribute), true)
                    .FirstOrDefault() as TableAttribute;
            var spec2 = targetType.GetCustomAttributes(typeof(PropertySetAttribute), true).Cast<PropertySetAttribute>().FirstOrDefault();
            if (spec == null && spec2 == null) return book.GetSheetAt(0);

            if (spec?.Index != null) return book.GetSheetAt((spec.Index.Value) - 1);
            if (spec2?.Index != null) return book.GetSheetAt((spec2.Index.Value) - 1);

            if (string.IsNullOrWhiteSpace(spec?.MatchPattern) && string.IsNullOrWhiteSpace(spec2?.MatchPattern)) return book.GetSheetAt(0);

            foreach (ISheet sheet in book)
            {
                if (Regex.IsMatch(sheet.SheetName, spec?.MatchPattern ?? spec2.MatchPattern, RegexOptions.IgnoreCase)) return sheet;
            }

            return null;
        }

        public static IEnumerable<Func<T, string, string, CellType, bool>> GetAssignFuncs<T>()
        {
            var list = new List<Func<T, string, string, CellType, bool>>();
            foreach (var prop in typeof(T).GetProperties())
            {
                var att = (ColumnAttribute) prop.GetCustomAttributes(typeof(ColumnAttribute), true).FirstOrDefault();
                if (att == null) continue;
                list.Add(GetAssignFunc<T>(att.MatchPattern, prop));
            }
            return list;
        }

        public static Func<T, string, string, CellType, bool> GetAssignFunc<T>(string keyName, PropertyInfo prop)
        {
            return (s, key, val, origType) =>
            {
                if (!Regex.IsMatch(key, keyName, RegexOptions.IgnoreCase)) return false;
                var converter = TypeDescriptor.GetConverter(prop.PropertyType);
                if (prop.PropertyType == typeof(DateTime))
                {
                    try
                    {
                        prop.SetValue(s, DateTime.FromOADate(double.Parse(val)));
                    }
                    catch (Exception e)
                    {
                        prop.SetValue(s, DateTime.Parse(val));
                    }
                }
                else if (prop.PropertyType == typeof(bool) || prop.PropertyType == typeof(bool?))
                {
                    if (new[] {"Y", "YES", "TRUE", "1"}.Any(x => x == val.ToUpper()))
                    {
                        prop.SetValue(s, true);
                    }
                    else if (!string.IsNullOrWhiteSpace(val))
                    {
                        prop.SetValue(s, false);
                    }
                }
                else
                {
                    prop.SetValue(s, converter.ConvertFromString(val));
                }
               return true;
            };
        }

        public static T WorksheetToModel<T>(ISheet sheet) where T : new()
        {
            // These are properties with types that themselves are mapped to the spreadsheet
            //var subProps = typeof(T).GetProperties().ToDictionary(x => x, y => y.PropertyType
            //    .GetCustomAttributes(typeof(TableAttribute), true)
            //    .Cast<Attribute>().FirstOrDefault()).Where(x => x.Value != null);

            var rows = sheet.GetRowEnumerator();
            var header = new Dictionary<int, string>();
            var model = new T();
            var populated = false;
            var handlers = GetAssignFuncs<T>();

            while (rows.MoveNext())
            {
                var row = (IRow)rows.Current;
                var prop = row.GetCell(0)?.StringCellValue;
                if (string.IsNullOrWhiteSpace(prop)) continue;
                var cell = row.GetCell(1);
                if (cell == null) continue;
                var origType = cell.CellType;
                cell.SetCellType(CellType.String);
                var value = cell.StringCellValue;
                if (string.IsNullOrWhiteSpace(value)) continue;

                if (handlers.Any(h => h(model, prop, value, origType)))
                {
                    populated = true;
                }
            }

            PopulateSubProperties(model, sheet.Workbook);
            
            return populated ? model : default(T);
        }

        private static void PopulateSubProperties<T>(T model, IWorkbook workbook)
        {
            foreach (var subProp in typeof(T).GetProperties())
            {
                var subType = subProp.PropertyType;
                if (typeof(IEnumerable).IsAssignableFrom(subType) && subType.GenericTypeArguments.Length > 0)
                    subType = subType.GetGenericArguments().FirstOrDefault();
                var tableAtt = subType.GetCustomAttributes(typeof(TableAttribute), true).Cast<TableAttribute>()
                    .FirstOrDefault();

                var subPropAtt = subType.GetCustomAttributes(typeof(PropertySetAttribute), true).Cast<PropertySetAttribute>()
                    .FirstOrDefault();

                if (tableAtt == null && subPropAtt == null) continue;
                
                var subSheet = GetSheet(workbook, subType);

                if (tableAtt != null)
                {
                    var models = (IEnumerable<dynamic>) ReflectionUtils.CallGeneric(
                        typeof(Porter).GetMethod("WorksheetToModels"), new[] {subType},
                        null, new object[] {subSheet, tableAtt.HeaderMatchPattern});

                    // We only filter if the attribute is true or absent
                    var att = (SubTableAttribute) subProp.GetCustomAttributes(typeof(SubTableAttribute))
                        .FirstOrDefault();
                    if (att == null || att.AutoFilter)
                    {
                        // Get all properties whose names have a matching property in the parent and (if any exist) restrict by matching values
                        var subKeys = subType.GetProperties()
                            .Where(x => typeof(T).GetProperties().Any(y => y.Name == x.Name)).ToArray();
                        if (subKeys.Any())
                        {
                            var keys = subKeys.ToDictionary(x => x, y => model.GetType().GetProperty(y.Name));
                            models = models.Where(x => keys.All(y => y.Key.GetValue(x) == y.Value.GetValue(model)))
                                .ToList();

                        }

                    }

                    var targetList = (IList) subProp.GetValue(model);
                    if (targetList == null)
                    {
                        subProp.SetValue(model, Activator.CreateInstance(subProp.PropertyType));
                        targetList = (IList) subProp.GetValue(model);
                    }
                    foreach (var m in models)
                    {
                        targetList.Add(m);
                    }

                }
                else
                {
                    var subModel = (dynamic)ReflectionUtils.CallGeneric(
                        typeof(Porter).GetMethod("WorksheetToModel"), new[] { subType },
                        null, new object[] { subSheet });
                    subProp.SetValue(model, subModel);
                }
            }
        }

        //public IEnumerable<T> WorksheetToModels<T>(ISheet sheet, string headerKey,
        //    Func<T, string, string, CellType, bool>[] handlers) where T : new()
        public static IEnumerable<T> WorksheetToModels<T>(ISheet sheet, string headerKey) where T : new()
        {
            var rows = sheet.GetRowEnumerator();
            var header = new Dictionary<int, string>();
            var handlers = GetAssignFuncs<T>();
            var models = new List<T>();

            while (rows.MoveNext())
            {
                var row = (IRow)rows.Current;
                if (string.Equals(row.GetCell(0)?.StringCellValue, headerKey, StringComparison.CurrentCultureIgnoreCase) && header.Keys.Count == 0)
                {
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        row.GetCell(i).SetCellType(CellType.String);
                        header.Add(i, row.GetCell(i).StringCellValue.ToUpper());
                    }
                }
                else if (header.Keys.Count > 0)
                {

                    var model = new T();
                    var populated = false;
                    for (var i = 0; i < row.LastCellNum; i++)
                    {
                        var origType = row.GetCell(i).CellType;
                        row.GetCell(i).SetCellType(CellType.String);
                        if (!header.ContainsKey(i)) continue;
                        var cell = row.GetCell(i);
                        if (cell == null || origType == CellType.Blank) continue;
                        var value = row.GetCell(i).StringCellValue;
                        
                        if (string.IsNullOrWhiteSpace(value)) continue;
                        var prop = header[i];

                        if (handlers.Any(h => h(model, prop, value, origType)))
                        {
                            populated = true;
                        }
                    }
                    if (populated)
                    {
                        models.Add(model);
                        PopulateSubProperties(model, sheet.Workbook);
                    }
                }
            }

            if (header.Values.Count == 0) throw new Exception($"No header row found in sheet {sheet.SheetName}.  Row should start with {headerKey} in first cell.");
            return models;
        }
    }
}
