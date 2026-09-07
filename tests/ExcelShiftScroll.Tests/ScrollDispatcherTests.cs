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

        Assert.True(dispatcher.Enqueue(3, new IntPtr(1)));
        Assert.True(dispatcher.Enqueue(2, new IntPtr(1)));
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
        dispatcher.Enqueue(3, new IntPtr(1));

        dispatcher.Dispose();
        queue.Actions[0]();

        Assert.Empty(scroller.Deltas);
        Assert.False(dispatcher.Enqueue(3, new IntPtr(1)));
    }

    private sealed class FakeMacroQueue : IExcelMacroQueue
    {
        internal List<Action> Actions { get; } = new();
        public void Queue(Action action) => Actions.Add(action);
    }

    [Fact]
    public void NewTargetDropsPreviousTargetsInput()
    {
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        using var dispatcher = new ScrollDispatcher(queue, scroller);
        dispatcher.Enqueue(3, new IntPtr(1));
        dispatcher.Enqueue(2, new IntPtr(2));
        queue.Actions[0]();
        Assert.Equal(new[] { 2 }, scroller.Deltas);
        Assert.Equal(new[] { new IntPtr(2) }, scroller.Targets);
    }

    [Fact]
    public void ExpiredBatchIsDropped()
    {
        long now = 0;
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        using var dispatcher = new ScrollDispatcher(queue, scroller, milliseconds: () => now);
        dispatcher.Enqueue(3, new IntPtr(1));
        now = 501;
        queue.Actions[0]();
        Assert.Empty(scroller.Deltas);
    }

    [Fact]
    public void FreshInputDoesNotReviveExpiredColumns()
    {
        long now = 0;
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        using var dispatcher = new ScrollDispatcher(queue, scroller, milliseconds: () => now);
        dispatcher.Enqueue(3, new IntPtr(1));
        now = 501;
        dispatcher.Enqueue(2, new IntPtr(1));
        queue.Actions[0]();
        Assert.Equal(new[] { 2 }, scroller.Deltas);
    }

    [Fact]
    public void DisabledBeforeDrainDropsInput()
    {
        var enabled = true;
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        using var dispatcher = new ScrollDispatcher(queue, scroller, isEnabled: () => enabled);
        dispatcher.Enqueue(3, new IntPtr(1));
        enabled = false;
        queue.Actions[0]();
        Assert.Empty(scroller.Deltas);
    }

    [Fact]
    public void CancelThenResumeDoesNotReplayOldInput()
    {
        var queue = new FakeMacroQueue();
        var scroller = new FakeScroller();
        using var dispatcher = new ScrollDispatcher(queue, scroller);
        dispatcher.Enqueue(3, new IntPtr(1));
        dispatcher.CancelPending();
        queue.Actions[0]();
        Assert.Empty(scroller.Deltas);
        dispatcher.Enqueue(2, new IntPtr(1));
        queue.Actions[1]();
        Assert.Equal(new[] { 2 }, scroller.Deltas);
    }

    [Fact]
    public void QueueFailurePassesThroughAndAllowsRetry()
    {
        var queue = new FailingMacroQueue();
        using var dispatcher = new ScrollDispatcher(queue, new FakeScroller());
        Assert.False(dispatcher.Enqueue(3, new IntPtr(1)));
        queue.Fail = false;
        Assert.True(dispatcher.Enqueue(2, new IntPtr(1)));
        Assert.False(dispatcher.Enqueue(2, IntPtr.Zero));
    }

    private sealed class FailingMacroQueue : IExcelMacroQueue
    {
        internal bool Fail = true;
        public void Queue(Action action)
        {
            if (Fail) { throw new InvalidOperationException(); }
        }
    }

    private sealed class FakeScroller : IExcelScroller
    {
        internal List<int> Deltas { get; } = new();
        internal List<IntPtr> Targets { get; } = new();
        public void ScrollColumns(int columnDelta, IntPtr targetWindow)
        {
            Deltas.Add(columnDelta);
            Targets.Add(targetWindow);
        }
    }
}
