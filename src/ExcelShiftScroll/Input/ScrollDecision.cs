namespace ExcelShiftScroll.Input;

public readonly struct ScrollDecision
{
    private ScrollDecision(bool handled, int columnDelta)
    {
        Handled = handled;
        ColumnDelta = columnDelta;
    }

    public bool Handled { get; }
    public int ColumnDelta { get; }

    public static ScrollDecision PassThrough => new(false, 0);
    public static ScrollDecision Consume(int columnDelta) => new(true, columnDelta);
}
