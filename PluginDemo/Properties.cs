using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CsvParameterUpdater
{
    internal static class Properties
    {
        /// <summary>
        /// Filters elements by category using the specified comparison operator.
        /// Supported operators:
        /// "equals", "not equal", "contains", "not contains", "exist".
        /// </summary>
        /// <param name="collector">Initial element collector</param>
        /// <param name="comparer">Comparison operator</param>
        /// <param name="value">Category name or condition</param>
        /// <returns>Filtered collection of elements</returns>

        internal static FilteredElementCollector FilterByCategory(FilteredElementCollector collector,
            Document doc, string comparer, string value)
        {
            switch (comparer.ToLower())
            {
                case "equals":
                    if (Enum.TryParse(value, out BuiltInCategory bicEQ))
                    {
                        return collector.OfCategory(bicEQ);
                    }
                    break;

                case "not equal":
                    if (Enum.TryParse(value, out BuiltInCategory bicNEQ))
                    {
                        var elementsOfCategory = new FilteredElementCollector(doc).OfCategory(bicNEQ).ToElementIds();
                        return collector.Excluding(elementsOfCategory);
                    }
                    break;

                case "contains":
                case "not contains":
                    bool shouldContain = comparer == "contains";
                    var elements = collector.ToList();

                    return new FilteredElementCollector(doc, elements
                        .Where(e => (e.Category?.Name?.Contains(value) == true) == shouldContain)
                        .Select(e => e.Id)
                        .ToList());

                case "exist":
                    var categoryParamId = new ElementId(BuiltInParameter.ELEM_CATEGORY_PARAM);
                    bool shouldExist = value.Equals("true", StringComparison.OrdinalIgnoreCase);

                    var rule = shouldExist
                        ? ParameterFilterRuleFactory.CreateHasValueParameterRule(categoryParamId)
                        : ParameterFilterRuleFactory.CreateHasNoValueParameterRule(categoryParamId);

                    return collector.WherePasses(new ElementParameterFilter(rule));
            }

            return collector;
        }

        /// <summary>
        /// Filters elements by family name using the specified comparison operator.
        /// Supported operators:
        /// "equals", "not equal", "contains", "not contains", "exist".
        /// </summary>
        /// <param name="collector">Initial element collector</param>
        /// <param name="comparer">Comparison operator</param>
        /// <param name="value">Family name or search keyword</param>
        /// <returns>Filtered collection of elements</returns>

        internal static FilteredElementCollector FilterByFamilyName(FilteredElementCollector collector,
            string comparer, string value)
        {
            var paramId = new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME);

            switch (comparer.ToLower())
            {
                case "equals":
                    return collector.WherePasses(new ElementParameterFilter(
                        ParameterFilterRuleFactory.CreateEqualsRule(paramId, value, false)));

                case "not equal":
                    return collector.WherePasses(new ElementParameterFilter(
                        ParameterFilterRuleFactory.CreateNotEqualsRule(paramId, value, false)));

                case "contains":
                    return collector.WherePasses(new ElementParameterFilter(
                        ParameterFilterRuleFactory.CreateContainsRule(paramId, value, false)));

                case "not contains":
                    var hasValueFilter = new ElementParameterFilter(
                        ParameterFilterRuleFactory.CreateHasValueParameterRule(paramId));

                    var containsFilter = new ElementParameterFilter(
                        ParameterFilterRuleFactory.CreateContainsRule(paramId, value, false));

                    var notContainsFilter = new LogicalAndFilter(
                        hasValueFilter,
                        new ElementParameterFilter(containsFilter.GetRules(), true));

                    return collector.WherePasses(notContainsFilter);

                case "exist":
                    bool shouldExist = value.Equals("true", StringComparison.OrdinalIgnoreCase);
                    var rule = shouldExist
                        ? ParameterFilterRuleFactory.CreateHasValueParameterRule(paramId)
                        : ParameterFilterRuleFactory.CreateHasNoValueParameterRule(paramId);

                    return collector.WherePasses(new ElementParameterFilter(rule));

                default:
                    return collector;
            }
        }

        /// <summary>
        /// Filters elements by type name (Element.Name) using the specified comparison operator.
        /// Supported operators:
        /// "equals", "not equal", "exist".
        /// </summary>
        /// <param name="collector">Initial element collector</param>
        /// <param name="comparer">Comparison operator</param>
        /// <param name="value">Type name or condition</param>
        /// <returns>Filtered collection of elements</returns>

        internal static FilteredElementCollector FilterByTypeName(FilteredElementCollector collector,
            Document doc, string comparer, string value)
        {
            var elements = collector.ToList();

            switch (comparer.ToLower())
            {
                case "equals":
                    return new FilteredElementCollector(doc, elements
                        .Where(e => e.Name?.Equals(value, StringComparison.OrdinalIgnoreCase) == true)
                        .Select(e => e.Id)
                        .ToList());

                case "not equal":
                    return new FilteredElementCollector(doc, elements
                        .Where(e => e.Name?.Equals(value, StringComparison.OrdinalIgnoreCase) != true)
                        .Select(e => e.Id)
                        .ToList());

                case "exist":
                    bool shouldExist = value.Equals("true", StringComparison.OrdinalIgnoreCase);
                    return new FilteredElementCollector(doc, elements
                        .Where(e => (e.Name != null) == shouldExist)
                        .Select(e => e.Id)
                        .ToList());

                default:
                    return collector;
            }
        }

        /// <summary>
        /// Filters elements by the name of their associated workset using the specified comparison operator.
        /// Supported operators:
        /// "equals", "not equal", "contains", "not contains", "exist".
        /// </summary>
        /// <param name="collector">Initial element collector</param>
        /// <param name="comparer">Comparison operator</param>
        /// <param name="value">Workset name or condition</param>
        /// <returns>Filtered collection of elements</returns>

        internal static FilteredElementCollector FilterByWorkset(FilteredElementCollector collector,
            Document doc, string comparer, string value)
        {
            var worksetTable = doc.GetWorksetTable();
            var elements = collector.ToList();

            switch (comparer.ToLower())
            {
                case "equals":
                    return new FilteredElementCollector(doc, elements
                        .Where(e => {
                            Workset workset = worksetTable.GetWorkset(e.WorksetId);
                            return workset?.Name?.Equals(value, StringComparison.OrdinalIgnoreCase) == true;
                        })
                        .Select(e => e.Id)
                        .ToList());

                case "not equal":
                    return new FilteredElementCollector(doc, elements
                        .Where(e => {
                            Workset workset = worksetTable.GetWorkset(e.WorksetId);
                            return workset?.Name?.Equals(value, StringComparison.OrdinalIgnoreCase) != true;
                        })
                        .Select(e => e.Id)
                        .ToList());

                case "contains":
                    return new FilteredElementCollector(doc, elements
                        .Where(e => {
                            Workset workset = worksetTable.GetWorkset(e.WorksetId);
                            return workset?.Name?.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
                        })
                        .Select(e => e.Id)
                        .ToList());

                case "not contains":
                    return new FilteredElementCollector(doc, elements
                        .Where(e => {
                            Workset workset = worksetTable.GetWorkset(e.WorksetId);
                            return workset?.Name?.IndexOf(value, StringComparison.OrdinalIgnoreCase) < 0;
                        })
                        .Select(e => e.Id)
                        .ToList());

                case "exist":
                    bool shouldExist = value.Equals("true", StringComparison.OrdinalIgnoreCase);
                    return new FilteredElementCollector(doc, elements
                        .Where(e => (worksetTable.GetWorkset(e.WorksetId) != null) == shouldExist)
                        .Select(e => e.Id)
                        .ToList());

                default:
                    return collector;
            }
        }
    }
}