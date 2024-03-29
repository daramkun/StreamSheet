using System.Text;

namespace Hazelnut.StreamSheet;

public class Cell
{
    public readonly int ColumnNumber;
    public readonly int RowNumber;

    public readonly int ColumnSpan;
    public readonly int RowSpan;

    public object? Value;

    public object? Style;

    public bool IsStringValue => Value is string;
    public bool IsFormulaValue => false;

    public Cell(int columnNumber, int rowNumber, object? value = null, int columnSpan = 1, int rowSpan = 1, object? style = null)
    {
        ColumnNumber = columnNumber;
        RowNumber = rowNumber;

        ColumnSpan = columnSpan;
        RowSpan = rowSpan;

        Value = value;
        Style = style;
    }

    public string ToAddress()
    {
        var builder = new StringBuilder();
        var col = ColumnNumber;

        while (col > 26)
        {
            int value = col % 26;
            if (value == 0)
            {
                col = col / 26 - 1;
                builder.Insert(0, 'A');
            }
            else
            {
                col /= 26;
                builder.Insert(0, (char)('A' + value - 1));
            }
        }

        if (col > 0)
            builder.Insert(0, (char)('A' + col - 1));

        builder.Append(RowNumber);

        return builder.ToString();
    }

    public override string ToString()
    {
        if (IsStringValue)
            return Value as string ?? string.Empty;
        return Value?.ToString() ?? string.Empty;
    }
}