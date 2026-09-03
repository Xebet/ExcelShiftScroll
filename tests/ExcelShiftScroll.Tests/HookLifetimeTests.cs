using System;
using ExcelShiftScroll.Input;
using Xunit;

namespace ExcelShiftScroll.Tests;

public sealed class HookLifetimeTests
{
    [Fact]
    public void RepeatedInstallAndDisposeDoNotDuplicateOrLeakHook()
    {
        var installs = 0;
        var uninstalls = 0;
        var lifetime = new HookLifetime(
            () =>
            {
                installs++;
                return new IntPtr(42);
            },
            handle =>
            {
                Assert.Equal(new IntPtr(42), handle);
                uninstalls++;
                return true;
            });

        lifetime.Install();
        lifetime.Install();
        lifetime.Dispose();
        lifetime.Dispose();

        Assert.Equal(1, installs);
        Assert.Equal(1, uninstalls);
        Assert.False(lifetime.IsInstalled);
    }

    [Fact]
    public void InstallFailureLeavesNoInstalledHook()
    {
        using var lifetime = new HookLifetime(() => IntPtr.Zero, handle => true);

        Assert.Throws<InvalidOperationException>(() => lifetime.Install());
        Assert.False(lifetime.IsInstalled);
    }
}
