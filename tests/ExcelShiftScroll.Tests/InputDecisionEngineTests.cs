using ExcelShiftScroll.Input;
using ExcelShiftScroll.Settings;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class InputDecisionEngineTests
{
    [Fact]
    public void ShiftWheelUpScrollsLeft()
    {
        var result = Decide(120);

        Assert.True(result.Handled);
        Assert.Equal(-3, result.ColumnDelta);
        Assert.Equal(-120, result.HorizontalWheelDelta);
    }

    [Fact]
    public void ShiftWheelDownScrollsRight()
    {
        var result = Decide(-120);

        Assert.True(result.Handled);
        Assert.Equal(3, result.ColumnDelta);
        Assert.Equal(120, result.HorizontalWheelDelta);
    }

    [Theory]
    [InlineData(120, 3)]
    [InlineData(-120, -3)]
    public void ReverseDirectionFlipsMapping(int wheelDelta, int expectedColumns)
    {
        var settings = ScrollSettings.Defaults();
        settings.ReverseDirection = true;

        var result = Decide(wheelDelta, settings);

        Assert.Equal(expectedColumns, result.ColumnDelta);
        Assert.Equal(wheelDelta, result.HorizontalWheelDelta);
    }

    [Fact]
    public void ConfiguredColumnsAreAppliedPerDetent()
    {
        var settings = ScrollSettings.Defaults();
        settings.ColumnsPerDetent = 5;

        Assert.Equal(-10, Decide(240, settings).ColumnDelta);
    }

    [Theory]
    [InlineData(ModifierKeys.None)]
    [InlineData(ModifierKeys.Shift | ModifierKeys.Control)]
    [InlineData(ModifierKeys.Shift | ModifierKeys.Alt)]
    [InlineData(ModifierKeys.Shift | ModifierKeys.Windows)]
    public void ModifierCombinationsOtherThanShiftAlonePassThrough(ModifierKeys modifiers)
    {
        var result = Decide(120, modifiers: modifiers);

        Assert.False(result.Handled);
    }

    [Fact]
    public void ExcelInBackgroundPassesThrough()
    {
        Assert.False(Decide(120, isExcelForeground: false).Handled);
    }

    [Fact]
    public void PointerOutsideWorksheetPassesThrough()
    {
        Assert.False(Decide(120, isWorksheetArea: false).Handled);
    }

    [Fact]
    public void NativeHorizontalWheelPassesThrough()
    {
        var input = new InputSnapshot(
            isVerticalWheel: false,
            isHorizontalWheel: true,
            wheelDelta: 120,
            modifiers: ModifierKeys.Shift,
            isExcelForeground: true,
            isWorksheetArea: true);

        Assert.False(new InputDecisionEngine().Decide(input, ScrollSettings.Defaults()).Handled);
    }

    [Fact]
    public void DisabledAddInPassesThrough()
    {
        var settings = ScrollSettings.Defaults();
        settings.Enabled = false;

        Assert.False(Decide(120, settings).Handled);
    }

    [Fact]
    public void HighResolutionDeltasAccumulateToOneDetent()
    {
        var engine = new InputDecisionEngine();
        var settings = ScrollSettings.Defaults();

        var first = engine.Decide(Snapshot(60), settings);
        var second = engine.Decide(Snapshot(60), settings);

        Assert.True(first.Handled);
        Assert.Equal(0, first.ColumnDelta);
        Assert.Equal(-60, first.HorizontalWheelDelta);
        Assert.Equal(-3, second.ColumnDelta);
        Assert.Equal(-60, second.HorizontalWheelDelta);
    }

    [Fact]
    public void IneligibleInputClearsPartialPrecisionDelta()
    {
        var engine = new InputDecisionEngine();
        var settings = ScrollSettings.Defaults();

        Assert.Equal(0, engine.Decide(Snapshot(60), settings).ColumnDelta);
        Assert.False(engine.Decide(Snapshot(120, ModifierKeys.None), settings).Handled);
        Assert.Equal(0, engine.Decide(Snapshot(60), settings).ColumnDelta);
    }

    [Fact]
    public void NativeDeliveryResetPreventsFallbackDoubleCounting()
    {
        var engine = new InputDecisionEngine();
        var settings = ScrollSettings.Defaults();
        engine.Decide(Snapshot(60), settings);
        engine.Reset();
        Assert.Equal(0, engine.Decide(Snapshot(60), settings).ColumnDelta);
        Assert.Equal(-3, engine.Decide(Snapshot(60), settings).ColumnDelta);
    }

    private static ScrollDecision Decide(
        int wheelDelta,
        ScrollSettings? settings = null,
        ModifierKeys modifiers = ModifierKeys.Shift,
        bool isExcelForeground = true,
        bool isWorksheetArea = true)
    {
        return new InputDecisionEngine().Decide(
            Snapshot(wheelDelta, modifiers, isExcelForeground, isWorksheetArea),
            settings ?? ScrollSettings.Defaults());
    }

    private static InputSnapshot Snapshot(
        int wheelDelta,
        ModifierKeys modifiers = ModifierKeys.Shift,
        bool isExcelForeground = true,
        bool isWorksheetArea = true) => new(
            isVerticalWheel: true,
            isHorizontalWheel: false,
            wheelDelta: wheelDelta,
            modifiers: modifiers,
            isExcelForeground: isExcelForeground,
            isWorksheetArea: isWorksheetArea);
}
