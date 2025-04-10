using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;

namespace CsvParameterUpdater
{
    

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        /// <summary>
        /// Main method for updating element parameters based on rules from an Excel (.xlsx) file.
        /// </summary>
        /// <remarks>
        /// Expected structure of the Excel file:
        /// - Rule: Name of the rule group (used to group conditions)
        /// - Property: Filtering property ("Category", "FamilyName", etc.)
        /// - Comparer: Comparison operator ("equals", "contains")
        /// - Value: Value to compare with
        /// - Parameter: Name of the parameter to update
        /// - ParameterValue: Value to assign to the parameter
        /// </remarks>

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var openDialog = new FileOpenDialog("Excel Files (*.xlsx)|*.xlsx");
            openDialog.Title = "Выберите Excel файл с правилами";

            if (openDialog.Show() != ItemSelectionDialogResult.Confirmed)
                return Result.Cancelled;

            string xlsxPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(openDialog.GetSelectedModelPath());

            TaskDialog.Show("Debug", $"Excel path: {xlsxPath}");

            if (string.IsNullOrEmpty(xlsxPath) || !File.Exists(xlsxPath))
            {
                TaskDialog.Show("Error", "Invalid or missing Excel file path.");
                return Result.Cancelled;
            }

            var rules = ReadXlsxToRuleDefinition(xlsxPath);

            if (rules.Count == 0)
                throw new InvalidOperationException("Excel файл не содержит правил");

            using (Transaction trans = new Transaction(doc, "Update Parameters from Rules"))
            {
                trans.Start();

                foreach (var rule in rules)
                {
                    var filteredElements = FilterElementsByConditions(doc, rule.Conditions);

                    if (filteredElements == null || !filteredElements.Any())
                    {
                        TaskDialog.Show("Информация",
                            $"Не найдены элементы, соответствующие правилу\n" +
                            $"Параметр: {rule.ParameterName}, Значение: {rule.ParameterValue}");
                        continue;
                    }

                    foreach (Element elem in filteredElements)
                    {
                        SetParameter(rule.ParameterName, rule.ParameterValue, elem);
                    }
                }

                trans.Commit();
            }

            TaskDialog.Show("Готово", "Параметры успешно обновлены по правилам");
            return Result.Succeeded;
        }

        /// <summary>
        /// Filters elements in the document based on the provided conditions.
        /// Conditions can include checks for category (Category), family name (FamilyName),
        /// and other properties using different comparison operators ("equals", "contains").
        /// </summary>
        /// <param name="conditions">
        /// List of filtering conditions where each condition includes:
        ///     Property - property to filter by ("Category", "FamilyName", etc.),
        ///     Comparer - comparison operator ("equals" for exact match, "contains" for partial match),
        ///     Value - value to compare with (e.g., category name, search string for FamilyName)
        /// </param>
        /// <returns>FilteredElementCollector containing the matching elements</returns>

        public FilteredElementCollector FilterElementsByConditions(Document doc, List<Condition> conditions)
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
                        collector = Properties.FilterByCategory(collector, doc, comparer, value);
                        break;

                    case "FamilyName":
                        collector = Properties.FilterByFamilyName(collector, comparer, value);
                        break;

                    case "TypeName":
                        collector = Properties.FilterByTypeName(collector, doc, comparer, value);
                        break;

                    case "Workset":
                        collector = Properties.FilterByWorkset(collector, doc, comparer, value);
                        break;

                    //case "TextParameterValue":
                    //    collector = Conditions.FilterByTextParameterValue(collector, doc, comparer, value, condition.ExtraInfo);
                    //    break;
                }
            }

            return collector;
        }


        /// <summary>
        /// Reads an XLSX file and converts it into a List<RuleDefinition>;
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <returns>List<RuleDefinition> with rule data extracted from the file</returns>

        private List<RuleDefinition> ReadXlsxToRuleDefinition(string filePath)
        {
            var rules = new List<RuleDefinition>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed();

                    if (rows.Count() < 2) return rules;

                    var firstRow = rows.First();
                    var ruleNames = new List<string>();
                    int col = 3;

                    while (col <= worksheet.LastColumnUsed().ColumnNumber())
                    {
                        var cell = firstRow.Cell(col);
                        if (!string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        {
                            ruleNames.Add(cell.Value.ToString());
                            col += 3;
                        }
                        else
                        {
                            break;
                        }
                    }


                    foreach (var row in rows.Skip(1))
                    {
                        string parameterName = row.Cell(1).Value.ToString().Trim();
                        string parameterValue = row.Cell(2).Value.ToString().Trim();

                        for (int ruleIndex = 0; ruleIndex < ruleNames.Count; ruleIndex++)
                        {
                            int valueIndex = 3 + (ruleIndex * 3);

                            string property = row.Cell(valueIndex).Value.ToString().Trim();
                            string comparer = row.Cell(valueIndex + 1).Value.ToString().Trim();
                            string value = row.Cell(valueIndex + 2).Value.ToString().Trim();

                            if (string.IsNullOrEmpty(property) ||
                                string.IsNullOrEmpty(comparer) ||
                                string.IsNullOrEmpty(value))
                                continue;

                            var rule = rules.FirstOrDefault(r =>
                                r.ParameterName == parameterName &&
                                r.ParameterValue == parameterValue);

                            if (rule == null)
                            {
                                rule = new RuleDefinition
                                {
                                    ParameterName = parameterName,
                                    ParameterValue = parameterValue,
                                    Conditions = new List<Condition>()
                                };
                                rules.Add(rule);
                            }

                            rule.Conditions.Add(new Condition
                            {
                                Property = property,
                                Comparer = comparer,
                                Value = value
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                TaskDialog.Show("Excel Error", $"Failed to read Excel file: {ex.Message}");
                return rules;
            }

            return rules;
        }

        /// <summary>
        /// Sets the value of an element's parameter with appropriate type conversion
        /// </summary>
        /// <param name="parameterName">Name of the parameter to set</param>
        /// <param name="value">Value to assign (will be converted to the proper type)</param>
        /// <param name="element">Element whose parameter will be modified</param>
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
    public class Condition
    {
        public string Property { get; set; }
        public string Comparer { get; set; }
        public string Value { get; set; }
    }
    public class RuleDefinition
    {
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }
        public List<Condition> Conditions { get; set; }
    }
    public class RuleGroup
    {
        public List<Condition> Conditions { get; set; }
        public string LogicalOperator { get; set; }
    }
}