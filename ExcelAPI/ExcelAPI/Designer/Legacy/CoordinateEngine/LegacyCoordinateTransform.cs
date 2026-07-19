using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.CoordinateEngine;

public class LegacyCoordinateTransform : ILegacyCoordinateTransform
{
    public string Name => "Identity";

    public CoordinateRect Transform(CoordinateRect rawNormalized, ClusterModel cluster, SheetModel sheet)
    {
        // Identity transform: pass through without modification.
        // When the real legacy transform is recovered from ConMasClient.exe,
        // replace this class with the correct implementation.
        return new CoordinateRect
        {
            Left = rawNormalized.Left,
            Top = rawNormalized.Top,
            Right = rawNormalized.Right,
            Bottom = rawNormalized.Bottom
        };
    }
}
