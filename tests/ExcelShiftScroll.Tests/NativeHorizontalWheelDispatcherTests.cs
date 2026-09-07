using System;
using System.Collections.Generic;
using ExcelShiftScroll.Excel;
using ExcelShiftScroll.Interop;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class NativeHorizontalWheelDispatcherTests
{
    [Fact]
    public void PostsUnmodifiedNativeHorizontalWheelWithSignedScreenPoint()
    {
        var messages = new List<PostedMessage>();
        var dispatcher = Create(3, messages);

        var result = dispatcher.TryPost(
            new IntPtr(42),
            new NativeMethods.Point(-20, 300),
            -120,
            3);

        Assert.True(result);
        var message = Assert.Single(messages);
        Assert.Equal(new IntPtr(42), message.Window);
        Assert.Equal((uint)NativeMethods.WmMouseHWheel, message.Message);
        Assert.Equal(-120, HighWordAsSignedShort(message.WParam));
        Assert.Equal(0, LowWord(message.WParam));
        Assert.Equal(-20, LowWordAsSignedShort(message.LParam));
        Assert.Equal(300, HighWordAsSignedShort(message.LParam));
    }

    [Fact]
    public void ScalesConfiguredColumnsAgainstWindowsHorizontalDistance()
    {
        var messages = new List<PostedMessage>();
        var dispatcher = Create(3, messages);

        Assert.True(dispatcher.TryPost(new IntPtr(1), default, -120, 1));
        Assert.Equal(-40, HighWordAsSignedShort(Assert.Single(messages).WParam));
    }

    [Fact]
    public void RetainsSubUnitPrecisionUntilItCanBePosted()
    {
        var messages = new List<PostedMessage>();
        var dispatcher = Create(3, messages);

        Assert.True(dispatcher.TryPost(new IntPtr(1), default, -1, 1));
        Assert.True(dispatcher.TryPost(new IntPtr(1), default, -1, 1));
        Assert.Empty(messages);

        Assert.True(dispatcher.TryPost(new IntPtr(1), default, -1, 1));
        Assert.Equal(-1, HighWordAsSignedShort(Assert.Single(messages).WParam));
    }

    [Theory]
    [InlineData(0u)]
    [InlineData(uint.MaxValue)]
    public void UnsupportedWindowsDistanceUsesComFallback(uint nativeColumns)
    {
        var messages = new List<PostedMessage>();
        var dispatcher = Create(nativeColumns, messages);

        Assert.False(dispatcher.TryPost(new IntPtr(1), default, -120, 3));
        Assert.Empty(messages);
    }

    [Fact]
    public void PostFailureRequestsComFallback()
    {
        var dispatcher = new NativeHorizontalWheelDispatcher(
            () => 3,
            (_, _, _, _) => false);

        Assert.False(dispatcher.TryPost(new IntPtr(1), default, -120, 3));
    }

    [Fact]
    public void ChangingTargetDiscardsFractionalRemainder()
    {
        var messages = new List<PostedMessage>();
        var dispatcher = Create(3, messages);
        dispatcher.TryPost(new IntPtr(1), default, -2, 1);
        dispatcher.TryPost(new IntPtr(2), default, -1, 1);
        Assert.Empty(messages);
        dispatcher.TryPost(new IntPtr(2), default, -2, 1);
        Assert.Equal(new IntPtr(2), Assert.Single(messages).Window);
    }

    [Fact]
    public void ResetDiscardsFractionalRemainder()
    {
        var messages = new List<PostedMessage>();
        var dispatcher = Create(3, messages);
        dispatcher.TryPost(new IntPtr(1), default, -2, 1);
        dispatcher.Reset();
        dispatcher.TryPost(new IntPtr(1), default, -1, 1);
        Assert.Empty(messages);
    }

    private static NativeHorizontalWheelDispatcher Create(
        uint nativeColumns,
        ICollection<PostedMessage> messages) =>
        new(
            () => nativeColumns,
            (window, message, wParam, lParam) =>
            {
                messages.Add(new PostedMessage(window, message, wParam, lParam));
                return true;
            });

    private static int LowWord(IntPtr value) => unchecked((ushort)value.ToInt64());

    private static int LowWordAsSignedShort(IntPtr value) =>
        unchecked((short)(ushort)value.ToInt64());

    private static int HighWordAsSignedShort(IntPtr value) =>
        unchecked((short)((ulong)value.ToInt64() >> 16));

    private readonly struct PostedMessage
    {
        internal PostedMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam)
        {
            Window = window;
            Message = message;
            WParam = wParam;
            LParam = lParam;
        }

        internal IntPtr Window { get; }
        internal uint Message { get; }
        internal IntPtr WParam { get; }
        internal IntPtr LParam { get; }
    }
}
