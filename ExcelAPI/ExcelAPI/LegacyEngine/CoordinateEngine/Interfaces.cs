using ExcelAPI.LegacyEngine.Models;

namespace ExcelAPI.LegacyEngine.CoordinateEngine;

public interface ICoordinateCalculator
{
    CoordinateRect CalculateRaw(ClusterModel cluster, SheetModel sheet);
    CoordinateRect CalculateNormalized(ClusterModel cluster, SheetModel sheet, double originX, double originY);
}

public interface ILegacyCoordinateTransform
{
    string Name { get; }
    CoordinateRect Transform(CoordinateRect rawNormalized, ClusterModel cluster, SheetModel sheet);
}
