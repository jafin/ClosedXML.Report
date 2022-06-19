using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests;

public class SubrangeGuardTest
{
    [Theory]
    [InlineData("A2:D5", "A3:D4", false, "All good")]
    [InlineData("A2:D5", "A2:D4", true, "Nested has same starting row as parent")]
    [InlineData("A2:D5", "A3:D5", true, "Nested has same ending row as parent")]
    // | {{Model.Name}} | Child               |                    |
    // | -------------- | ------------------- | ------------------ |
    // |                | {{item.ParentName}} |
    // |                |                     | {{item.ChildName}} |
    public void OverlappingChildRangeWithThrowException(string parentRange, string nestedRange, bool shouldThrowException, string reason)
    {
        var model = new ParentsModel
        {
            Name = "ParentsModel",
            Parents = new List<Parent>
            {
                new() { Name = "ParentA" },
                new() { Name = "ParentB" },
            }
        };

        var wbTemplate = new XLWorkbook();
        var ws1 = wbTemplate.AddWorksheet();

        ws1.Cell("A1").Value = "{{Model.Name}}";
        ws1.Cell("B1").Value = "Parent";
        ws1.Cell("C1").Value = "Child";
        ws1.Cell("B2").Value = "{{item.ParentName}}";
        ws1.Cell("C3").Value = "{{item.ChildName}}";
        var template = new XLTemplate(wbTemplate);

        var ws = template.Workbook.Worksheets.First();

        ws.Range(parentRange).AddToNamed("Model_Parents");
        ws.Range(nestedRange).AddToNamed("Model_Parents_Children");

        template.AddVariable("Model", model);

        Action act = () => template.Generate();

        if (shouldThrowException)
        {
            act.Should().Throw<InvalidNestedRangeException>(reason);
        }
        else
        {
            act.Should().NotThrow(reason);
        }
    }
}
