using System.Runtime.Serialization;

namespace ExcelShiftScroll.Settings;

[DataContract]
public sealed class ScrollSettings
{
    public static readonly int[] AllowedColumnCounts = { 1, 2, 3, 5, 10 };

    [DataMember(Name = "enabled", Order = 1)]
    public bool Enabled { get; set; } = true;

    [DataMember(Name = "columnsPerDetent", Order = 2)]
    public int ColumnsPerDetent { get; set; } = 3;

    [DataMember(Name = "reverseDirection", Order = 3)]
    public bool ReverseDirection { get; set; }

    [DataMember(Name = "diagnosticsEnabled", Order = 4)]
    public bool DiagnosticsEnabled { get; set; }

    public static ScrollSettings Defaults() => new();

    [OnDeserializing]
    private void InitializeDefaults(StreamingContext context)
    {
        Enabled = true;
        ColumnsPerDetent = 3;
    }

    public ScrollSettings ValidatedCopy()
    {
        var columns = System.Array.IndexOf(AllowedColumnCounts, ColumnsPerDetent) >= 0
            ? ColumnsPerDetent
            : 3;

        return new ScrollSettings
        {
            Enabled = Enabled,
            ColumnsPerDetent = columns,
            ReverseDirection = ReverseDirection,
            DiagnosticsEnabled = DiagnosticsEnabled,
        };
    }

    public ScrollSettings Copy() => new()
    {
        Enabled = Enabled,
        ColumnsPerDetent = ColumnsPerDetent,
        ReverseDirection = ReverseDirection,
        DiagnosticsEnabled = DiagnosticsEnabled,
    };
}
