using System;
using System.Collections.Generic;
using ExcelShiftScroll.Excel;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class ScrollDispatcherTests
{
    [Fact]
    public void RapidInputsAreCombinedIntoOneExcelMacro()
    {
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        using var dispatcher = new ScrollDispatcher(queue, scroller);

        Assert.True(dispatcher.Enqueue(3));
        Assert.True(dispatcher.Enqueue(2));
        Assert.Single(queue.Actions);

        queue.Actions[0]();

        Assert.Equal(new[] { 5 }, scroller.Deltas);
    }

    [Fact]
    public void DisposeDropsPendingWorkAndRejectsNewInput()
    {
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        var dispatcher = new ScrollDispatcher(queue, scroller);
        dispatcher.Enqueue(3);

        dispatcher.Dispose();
        queue.Actions[0]();

        Assert.Empty(scroller.Deltas);
        Assert.False(dispatcher.Enqueue(3));
    }

    private sealed class FakeMacroQueue : IExcelMacroQueue
    {
        internal List<Action> Actions { get; } = new();
        public void Queue(Action action) => Actions.Add(action);
    }

    private sealed class FakeScroller : IExcelScroller
    {
        internal List<int> Deltas { get; } = new();
        public void ScrollColumns(int columnDelta) => Deltas.Add(columnDelta);
    }
}
