using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using Bogus.DataSets;
using Xunit;

// ReSharper disable InconsistentNaming

namespace ICG.NetCore.Utilities.Spreadsheet.Tests;

public class TypeDiscovererTests
{
    [Fact]
    public void Sets_DisplayName_From_Annotation_If_Present()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_Annotation_If_Present_TestCase));

        //results.Should().HaveCount(1);
        Assert.Equal(1, results.Count);

        //results.First().DisplayName.Should().Be("Some Prop Name");
        Assert.Equal("Some Prop Name", results.First().DisplayName);
    }

    [Fact]
    public void Sets_DisplayName_From_Property_Name_If_No_Annotation()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_Property_Name_If_No_Annotation_TestCase));

        //results.Should().HaveCount(1);
        Assert.Equal(1, results.Count);

        //results.First().DisplayName.Should().Be("Some Prop");
        Assert.Equal("Some Prop", results.First().DisplayName);

    }

    [Fact]
    public void Sets_DisplayName_From_SpreadsheetColumn_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_SpreadsheetColumn_Attribute_TestCase));

        //results.Should().HaveCount(1);
        Assert.Equal(1, results.Count);


        //results.First().DisplayName.Should().Be("Some Prop Name");
        Assert.Equal("Some Prop Name", results.First().DisplayName);

    }

    [Fact]
    public void Sets_DisplayName_From_Display_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Sets_DisplayName_From_DisplayAttribute_Attribute_TestCase));

        //results.Should().HaveCount(1);
        Assert.Equal(1, results.Count);

        //results.First().DisplayName.Should().Be("Some Prop Name");
        Assert.Equal("Some Prop Name", results.First().DisplayName);

    }

    [Fact]
    public void Property_Excluded_From_SpreadsheetIgnore_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Property_Excluded_From_SpreadsheetColumn_Attribute_TestCase));

        //results.Should().HaveCount(1);
        Assert.Equal(1, results.Count);


        //results.Should().NotContain(d => d.DisplayName == "Ignored");
        Assert.False(results.Any(d => d.DisplayName == "Ignored"));

        //results.Should().Contain(d => d.DisplayName == "Real Column");
        Assert.True(results.Any(d => d.DisplayName == "Real Column"));

    }

    [Fact]
    public void Width_Is_Set_From_SpreadsheetColumn_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Width_Is_Set_From_SpreadsheetColumn_Attribute_TestCase));
        //results.First().Width.Should().Be(100);
        Assert.Equal(100, results.First().Width);

    }

    [Fact]
    public void Format_Is_Set_From_SpreadsheetColumn_Attribute()
    {
        var results = TypeDiscoverer.GetProps(typeof(Format_Is_Set_From_SpreadsheetColumn_Attribute_TestCase));
        //results.First().Format.Should().Be("c");
        Assert.Equal("c", results.First().Format);

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