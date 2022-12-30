using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using Bogus.DataSets;
using FluentAssertions;
using Xunit;

// ReSharper disable InconsistentNaming

namespace ICG.NetCore.Utilities.Spreadsheet.Tests;

public class TypeDiscovererTests
{
    [Fact]
    public void Sets_DisplayName_From_Annotation_If_Present()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_Annotation_If_Present_TestCase));

        results.Should().HaveCount(1);

        results.First().DisplayName.Should().Be("Some Prop Name");
    }

    [Fact]
    public void Sets_DisplayName_From_Property_Name_If_No_Annotation()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_Property_Name_If_No_Annotation_TestCase));

        results.Should().HaveCount(1);

        results.First().DisplayName.Should().Be("Some Prop");
    }

    [Fact]
    public void Sets_DisplayName_From_SpreadsheetColumn_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_SpreadsheetColumn_Attribute_TestCase));

        results.Should().HaveCount(1);

        results.First().DisplayName.Should().Be("Some Prop Name");
    }

    [Fact]
    public void Sets_DisplayName_From_Display_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_DisplayAttribute_Attribute_TestCase));

        results.Should().HaveCount(1);

        results.First().DisplayName.Should().Be("Some Prop Name");
    }

    [Fact]
    public void Property_Excluded_From_SpreadsheetIgnore_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Property_Excluded_From_SpreadsheetColumn_Attribute_TestCase));

        results.Should().HaveCount(1);

        results.Should().NotContain(d => d.DisplayName == "Ignored");
        results.Should().Contain(d => d.DisplayName == "Real Column");
    }

    [Fact]
    public void Width_Is_Set_From_SpreadsheetColumn_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Width_Is_Set_From_SpreadsheetColumn_Attribute_TestCase));
        results.First().Width.Should().Be(100);
    }

    [Fact]
    public void Format_Is_Set_From_SpreadsheetColumn_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Format_Is_Set_From_SpreadsheetColumn_Attribute_TestCase));
        results.First().Format.Should().Be("c");
    }

    private class Sets_DisplayName_From_Annotation_If_Present_TestCase
    {
        [DisplayName("Some Prop Name")]
        public string SomeProp { get; set; }
    }

    private class Sets_DisplayName_From_Property_Name_If_No_Annotation_TestCase
    {
        public string SomeProp { get; set; }
    }

    private class Sets_DisplayName_From_SpreadsheetColumn_Attribute_TestCase
    {
        [SpreadsheetColumn("Some Prop Name")]
        public string SomeProp { get; set; }
    }

    private class Sets_DisplayName_From_DisplayAttribute_Attribute_TestCase
    {
        [Display(Name = "Some Prop Name")]
        public string SomeProp { get; set; }
    }

    private class Property_Excluded_From_SpreadsheetColumn_Attribute_TestCase
    {
        [SpreadsheetColumn(ignore: true)]
        public string Ignored { get; set; }

        public string RealColumn { get; set; }
    }

    private class Width_Is_Set_From_SpreadsheetColumn_Attribute_TestCase
    {
        [SpreadsheetColumn(width: 100)]
        public string Column { get; set; }
    }

    private class Format_Is_Set_From_SpreadsheetColumn_Attribute_TestCase
    {
        [SpreadsheetColumn(format: "c")]
        public string Column { get; set; }
    }


}