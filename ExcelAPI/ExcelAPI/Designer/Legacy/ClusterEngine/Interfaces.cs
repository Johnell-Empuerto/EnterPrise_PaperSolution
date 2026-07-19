using ExcelAPI.Designer.Legacy.Models;

namespace ExcelAPI.Designer.Legacy.ClusterEngine;

public interface IClusterDetector
{
    List<ClusterModel> Detect(SheetModel sheet);
}

public interface IClusterBuilder
{
    ClusterModel Build(CellModel cell, SheetModel sheet, int clusterId, int sheetIndex);
    List<ClusterModel> BuildAll(SheetModel sheet, int sheetIndex);
}
