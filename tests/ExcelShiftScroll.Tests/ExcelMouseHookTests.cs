using System;
using ExcelShiftScroll.Input;
using ExcelShiftScroll.Interop;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class ExcelMouseHookTests
{
    [Theory]
    [InlineData(-1, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(3, false)]
    public void OnlyActionConsumesVerticalWheel(int code, bool expected)
    {
        Assert.Equal(expected, ExcelMouseHook.ShouldProcess(code, new IntPtr(NativeMethods.WmMouseWheel)));
    }

    [Fact]
    public void NativeHorizontalWheelPassesThrough()
    {
        Assert.False(ExcelMouseHook.ShouldProcess(0, new IntPtr(NativeMethods.WmMouseHWheel)));
    }
}
