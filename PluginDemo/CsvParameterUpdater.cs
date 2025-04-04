using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;

namespace CsvParameterUpdater
{
    public class Condition
    {
        public string Property { get; set; }
        public string Comparer { get; set; }
        public string Value { get; set; }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        /// <summary>
        /// Основной метод для обновления параметров элементов по правилам из CSV-файла
        /// </summary>
        /// <remarks>
        /// Ожидаемая структура CSV-файла:
        /// - Rule: Имя группы правил (группировка условий)
        /// - Property: Свойство для фильтрации ("Category", "FamilyName")
        /// - Comparer: Оператор сравнения ("equals", "contains")
        /// - Value: Значение для сравнения
        /// - Parameter: Имя параметра для обновления
        /// - ParameterValue: Значение параметра для установки
        /// </remarks>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var openDialog = new FileOpenDialog("CSV Files (*.csv)|*.csv");
            openDialog.Title = "Выберите CSV файл с правилами";

            if (openDialog.Show() != ItemSelectionDialogResult.Confirmed)
                return Result.Cancelled;

            string csvPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(openDialog.GetSelectedModelPath());

            DataTable rulesTable = ReadCsvToDataTable(csvPath);

            if (rulesTable.Rows.Count == 0)
                throw new InvalidOperationException("CSV файл не содержит правил");

            using (Transaction trans = new Transaction(doc, "Update Parameters from Rules"))
            {
                trans.Start();

                var ruleGroups = rulesTable.AsEnumerable()
                    .GroupBy(row => row.Field<string>("Rule"));

                foreach (var ruleGroup in ruleGroups)
                {
                    var conditions = ruleGroup.Select(row => new Condition
                    {
                        Property = row.Field<string>("Property"),
                        Comparer = row.Field<string>("Comparer"),
                        Value = row.Field<string>("Value")
                    }).ToList();

                    var parameterUpdate = ruleGroup.First();
                    string paramName = parameterUpdate.Field<string>("Parameter");
                    string paramValue = parameterUpdate.Field<string>("ParameterValue");

                    var filteredElements = FilterElementsByConditions(doc, conditions);

                    if (filteredElements == null || !filteredElements.Any())
                    {
                        TaskDialog.Show("Информация",
                            $"Не найдены элементы, соответствующие правилу: {ruleGroup.Key}\n" +
                            $"Параметр: {paramName}, Значение: {paramValue}");
                        continue;
                    }

                    foreach (Element elem in filteredElements)
                    {
                        SetParameter(paramName, paramValue, elem);
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Готово", "Параметры успешно обновлены по правилам");
            return Result.Succeeded;
        }

        /// <summary>
        /// Фильтрует элементы в документе на основе заданных условий
        /// Условия могут включать проверку категории (Category) или имени семейства (FamilyName)
        /// с различными операторами сравнения (equals, contains)
        /// </summary>
        /// <param name="conditions">Список условий для фильтрации, где каждое условие содержит:
        ///     Property - свойство для фильтрации ("Category" или "FamilyName"),
        ///     Comparer - оператор сравнения ("equals" для точного определения, "contains" для частичного совпадения),
        ///     Value - значение для сравнения (для Category - имя встроенной категории, для FamilyName - строка поиска)
        /// </param>
        /// <returns>Коллектор элементов с отфильтрованными элементами</returns>
        private FilteredElementCollector FilterElementsByConditions(Document doc, List<Condition> conditions)
        {
            var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();

            foreach (var condition in conditions)
            {
                string property = condition.Property;
                string comparer = condition.Comparer;
                string value = condition.Value;

                switch (property)
                {
                    case "Category":
                        collector = FilterByCategory(collector, doc, comparer, value);
                        break;

                    case "FamilyName":
                        collector = FilterByFamilyName(collector, comparer, value);
                        break;
                }
            }

            return collector;
        }

        /// <summary>
        /// Фильтрует элементы по категории с использованием указанного оператора сравнения
        /// </summary>
        /// <param name="comparer">Оператор сравнения:
        ///     "equals" - точное совпадение с встроенной категорией
        ///     "contains" - частичное совпадение с встроенной категорией</param>
        /// <param name="value">Значение для сравнения:</param>
        /// <returns>Коллектор элементов, отфильтрованный по категории</returns>
        /// <remarks>
        /// В текущей реализации поддерживается только оператор "equals" для встроенных категорий.
        /// </remarks>
        private FilteredElementCollector FilterByCategory(FilteredElementCollector collector,
            Document doc, string comparer, string value)
        {
            if (comparer == "equals")
            {
                if (Enum.TryParse(value, out BuiltInCategory bic))
                {
                    return collector.OfCategory(bic);
                }
            }
            //else if (comparer == "contains")
            //{
            //    var elements = collector.ToList();
            //    return new FilteredElementCollector(doc, elements
            //        .Where(e => e.Category?.Name?.Contains(value) == true)
            //        .Select(e => e.Id)
            //        .ToList());
            //}

            return collector;
        }

        /// <summary>
        /// Фильтрует элементы по имени семейства с использованием указанного оператора сравнения
        /// </summary>
        /// <param name="comparer">Оператор сравнения:
        ///     "equals" - точное совпадение имени семейства,
        ///     "contains" - частичное совпадение имени семейства</param>
        /// <param name="value">Значение для сравнения (строка поиска для имени семейства)</param>
        /// <returns>Коллектор элементов, отфильтрованный по имени семейства</returns>
        /// <remarks>
        /// Использует встроенный параметр ALL_MODEL_FAMILY_NAME для фильтрации.
        /// Регистр символов не учитывается (case insensitive).
        /// Для "equals" выполняется точное сравнение, для "contains" - поиск подстроки.
        /// </remarks>
        private FilteredElementCollector FilterByFamilyName(FilteredElementCollector collector,
            string comparer, string value)
        {
            if (comparer == "equals")
            {
                var rule = ParameterFilterRuleFactory.CreateEqualsRule(
                    new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME),
                    value,
                    false);
                return collector.WherePasses(new ElementParameterFilter(rule));
            }
            else if (comparer == "contains")
            {
                var rule = ParameterFilterRuleFactory.CreateContainsRule(
                    new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME),
                    value,
                    false);
                return collector.WherePasses(new ElementParameterFilter(rule));
            }

            return collector;
        }

        /// <summary>
        /// Читает CSV-файл и преобразует его в DataTable
        /// </summary>
        /// <param name="filePath">Путь к CSV-файлу</param>
        /// <returns>DataTable с данными из CSV</returns>
        private DataTable ReadCsvToDataTable(string filePath)
        {
            DataTable dt = new DataTable();

            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] headers = sr.ReadLine().Split(';');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header.Trim('"').Trim());
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(';');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length && i < rows.Length; i++)
                    {
                        dr[i] = rows[i].Trim('"').Trim();
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        /// <summary>
        /// Устанавливает значение параметра элемента с преобразованием типов (из Select)
        /// </summary>
        /// <param name="parameterName">Имя параметра для установки</param>
        /// <param name="value">Значение для установки (будет преобразовано в соответствующий тип)</param>
        /// <param name="element">Элемент, параметр которого нужно изменить</param>
        public static void SetParameter(string parameterName, object value, Element element)
        {
            try
            {
                Parameter p = element.LookupParameter(parameterName);
                if (p == null) return;
                if (p.IsReadOnly) return;

                if (value is bool && p.StorageType == StorageType.Integer)
                {
                    p.Set((bool)value ? 1 : 0);
                }
                else if (value is int && p.StorageType == StorageType.Integer)
                {
                    p.Set((int)value);
                }
                else if (value is double && p.StorageType == StorageType.Double)
                {
                    p.Set((double)value);
                }
                else if (value is int && p.StorageType == StorageType.ElementId)
                {
                    p.Set(new ElementId((int)value));
                }
                else if (value is ElementId && p.StorageType == StorageType.ElementId)
                {
                    p.Set((ElementId)value);
                }
                else if (value is ElementId && p.StorageType == StorageType.Integer)
                {
                    p.Set(((ElementId)value).IntegerValue);
                }
                else if (value is ElementId && p.StorageType == StorageType.String)
                {
                    p.Set(((ElementId)value).ToString());
                }
                else if (value is string && p.StorageType == StorageType.String)
                {
                    p.Set((string)value);
                }
                else
                {
                    TaskDialog.Show("Error", $"Parameter and value were incompatible.\nParameter: {parameterName}\nValue: {value}");
                }
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", $"An error occurred setting parameter {parameterName} to a value of {value}.\n Error message: {e.Message}");
            }
        }
    }
}