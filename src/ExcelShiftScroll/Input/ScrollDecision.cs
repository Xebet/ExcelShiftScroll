namespace ExcelShiftScroll.Input;

public readonly struct ScrollDecision
{
    private ScrollDecision(bool handled, int columnDelta, int horizontalWheelDelta)
    {
        Handled = handled;
        ColumnDelta = columnDelta;
        HorizontalWheelDelta = horizontalWheelDelta;
    }

    public bool Handled { get; }
    public int ColumnDelta { get; }
    public int HorizontalWheelDelta { get; }

    public static ScrollDecision PassThrough => new(false, 0, 0);
    public static ScrollDecision Consume(int columnDelta, int horizontalWheelDelta) =>
        new(true, columnDelta, horizontalWheelDelta);
}
