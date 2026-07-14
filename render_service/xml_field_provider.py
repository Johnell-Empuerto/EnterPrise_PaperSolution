import sys, os
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent
sys.path.insert(0, str(_PROJECT_ROOT))

import psycopg2
from CoordinateEngine.models.field_def import FieldDef
from CoordinateEngine.utils.file_utils import parse_range
from render_service.db_config import DB_CONFIG


def get_clusters(template_id: int):
    query = """
        SELECT ds.def_sheet_no,
               dc.cluster_id,
               dc.cluster_name,
               dc.cluster_type,
               dc.cell_addr,
               dc.input_parameter,
               dc.left_position,
               dc.top_position,
               dc.right_position,
               dc.bottom_position
        FROM def_sheet ds
        JOIN def_cluster dc ON ds.def_sheet_id = dc.def_sheet_id
        WHERE ds.def_top_id = %s
        ORDER BY ds.def_sheet_no, dc.cluster_id
    """
    conn = psycopg2.connect(**DB_CONFIG)
    try:
        with conn.cursor() as cur:
            cur.execute(query, (template_id,))
            rows = cur.fetchall()
    finally:
        conn.close()

    clusters = []
    for row in rows:
        sheet_no, cluster_id, name, ctype, cell_addr, input_param, left_pos, top_pos, right_pos, bottom_pos = row
        clusters.append({
            "sheet_no": sheet_no,
            "cluster_id": cluster_id,
            "name": name,
            "type": ctype,
            "cell_addr": cell_addr,
            "input_parameter": input_param,
            "left_position": left_pos,
            "top_position": top_pos,
            "right_position": right_pos,
            "bottom_position": bottom_pos,
        })
    return clusters


def cluster_to_fielddef(cluster: dict) -> FieldDef:
    cell_addr = cluster["cell_addr"]
    if not cell_addr:
        raise ValueError(f"Cluster {cluster['cluster_id']} has no cell_addr")
    parsed = parse_range(cell_addr)
    if parsed is None:
        raise ValueError(f"Could not parse cell_addr: {cell_addr}")
    col, row, col_end, row_end = parsed

    input_param = cluster.get("input_parameter") or ""
    required = False
    data_type = cluster.get("type") or "text"
    if "Required=1" in input_param:
        required = True
    if "Restriction=Number" in input_param or "Restriction=Integer" in input_param:
        data_type = "number"
    elif "Restriction=Date" in input_param:
        data_type = "date"
    elif "Restriction=Email" in input_param:
        data_type = "email"

    sheet_no = cluster.get("sheet_no", 1)
    sheet_index = max(0, sheet_no - 1)

    def to_float(v):
        if v is None: return None
        try: return float(v)
        except: return None

    ratio_left = to_float(cluster.get("left_position"))
    ratio_top = to_float(cluster.get("top_position"))
    ratio_right = to_float(cluster.get("right_position"))
    ratio_bottom = to_float(cluster.get("bottom_position"))

    return FieldDef(
        addr=cell_addr,
        col=col,
        row=row,
        col_end=col_end,
        row_end=row_end,
        sheet_index=sheet_index,
        value=None,
        is_merge=True if (col != col_end or row != row_end) else False,
        data_type=data_type,
        required=required,
        ratio_left=ratio_left,
        ratio_top=ratio_top,
        ratio_right=ratio_right,
        ratio_bottom=ratio_bottom,
    )


def get_cluster_fielddefs(template_id: int):
    clusters = get_clusters(template_id)
    return [cluster_to_fielddef(c) for c in clusters], clusters
