using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelShiftScroll.AddIn;
using ExcelShiftScroll.Settings;

namespace ExcelShiftScroll.Ribbon;

[ComVisible(true)]
[Guid("4BE5B85E-38AD-4202-9979-CC7482903462")]
public sealed class ExcelShiftScrollRibbon : ExcelRibbon
{
    private IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab idMso='TabAddIns'>
        <group id='ExcelShiftScroll.Group' label='Shift Scroll'>
          <toggleButton id='ExcelShiftScroll.Enabled'
                        label='Enabled'
                        screentip='Enable Shift + wheel horizontal scrolling'
                        getPressed='GetEnabled'
                        onAction='OnEnabledChanged'/>
          <dropDown id='ExcelShiftScroll.Columns'
                    label='Columns per detent'
                    getSelectedItemIndex='GetSelectedColumnIndex'
                    onAction='OnColumnsChanged'>
            <item id='ExcelShiftScroll.Columns1' label='1'/>
            <item id='ExcelShiftScroll.Columns2' label='2'/>
            <item id='ExcelShiftScroll.Columns3' label='3'/>
            <item id='ExcelShiftScroll.Columns5' label='5'/>
            <item id='ExcelShiftScroll.Columns10' label='10'/>
          </dropDown>
          <checkBox id='ExcelShiftScroll.Reverse'
                    label='Reverse direction'
                    getPressed='GetReverseDirection'
                    onAction='OnReverseDirectionChanged'/>
          <button id='ExcelShiftScroll.Reset'
                  label='Restore defaults'
                  onAction='OnRestoreDefaults'/>
          <button id='ExcelShiftScroll.About'
                  label='About'
                  onAction='OnAbout'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

    public void OnLoad(IRibbonUI ribbonUi)
    {
        _ribbonUi = ribbonUi;
    }

    public bool GetEnabled(IRibbonControl control) => AppServices.Settings.Current.Enabled;

    public bool GetReverseDirection(IRibbonControl control) =>
        AppServices.Settings.Current.ReverseDirection;

    public int GetSelectedColumnIndex(IRibbonControl control)
    {
        var columns = AppServices.Settings.Current.ColumnsPerDetent;
        var index = Array.IndexOf(ScrollSettings.AllowedColumnCounts, columns);
        return index >= 0 ? index : 2;
    }

    public void OnEnabledChanged(IRibbonControl control, bool pressed)
    {
        Update(settings => settings.Enabled = pressed);
    }

    public void OnReverseDirectionChanged(IRibbonControl control, bool pressed)
    {
        Update(settings => settings.ReverseDirection = pressed);
    }

    public void OnColumnsChanged(IRibbonControl control, string selectedId, int selectedIndex)
    {
        if (selectedIndex < 0 || selectedIndex >= ScrollSettings.AllowedColumnCounts.Length)
        {
            return;
        }

        var columns = ScrollSettings.AllowedColumnCounts[selectedIndex];
        Update(settings => settings.ColumnsPerDetent = columns);
    }

    public void OnRestoreDefaults(IRibbonControl control)
    {
        AppServices.Settings.Reset();
        _ribbonUi?.Invalidate();
    }

    public void OnAbout(IRibbonControl control)
    {
        var settings = AppServices.Settings.Current;
        var message = string.Format(
            CultureInfo.CurrentCulture,
            "ExcelShiftScroll {1}{0}{0}Status: {2}{0}Enabled: {3}{0}Columns per detent: {4}{0}Reverse direction: {5}{0}{0}No network access, telemetry, or workbook-content collection.",
            Environment.NewLine,
            VersionFunction.Version(),
            AppServices.Status,
            settings.Enabled,
            settings.ColumnsPerDetent,
            settings.ReverseDirection);

        MessageBox.Show(
            message,
            "About ExcelShiftScroll",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void Update(Action<ScrollSettings> update)
    {
        AppServices.Settings.Update(update);
        _ribbonUi?.Invalidate();
    }
}
