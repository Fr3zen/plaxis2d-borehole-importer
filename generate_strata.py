# 【模块中文说明】
# 离线地层几何生成器。
# 输入为统一层序的钻孔 JSON：x 为沿剖面距离，ground/bottoms 为标高，长度单位为 m。
# 逐相邻孔区间直线连接同名层面，允许层厚为零及尖灭后再次出现，不进行端孔外推。
# 输出验证报告、GeoJSON、剖面预览及可审阅的 PLAXIS 几何命令；本模块不连接软件、不赋材、不计算。
# 
"""Generate validated 2D strata from borehole interface elevations, offline."""
import argparse
import json
import math
from pathlib import Path

from shapely.geometry import Polygon, mapping
from shapely.ops import unary_union


# 输入统一层序的地层字典，返回排序钻孔、模型域、多层区域和几何验证报告。
# 相邻孔之间分段拼接可以保留尖灭及再出现；完全零厚度层保留逻辑编号但没有面积。
# 最后检查各层之间的重叠、空洞及超界面积；几何或数值不合法时抛 ValueError。
def build_model(data):
    if data.get("units") != "m":
        raise ValueError("This version requires units='m' and elevations in metres.")
    layers = data["layers"]
    holes = sorted(data["boreholes"], key=lambda item: item["x"])
    if not layers or len(holes) < 2:
        raise ValueError("At least one layer and two boreholes are required.")
    if len({layer["id"] for layer in layers}) != len(layers):
        raise ValueError("Layer identifiers must be unique.")
    if len({hole["id"] for hole in holes}) != len(holes):
        raise ValueError("Borehole identifiers must be unique.")
    previous_x = None
    for hole in holes:
        values = [hole["x"], hole["ground"], *hole["bottoms"]]
        if any(isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) for v in values):
            raise ValueError(f"{hole['id']}: missing/non-finite values; pinch-outs require explicit design.")
        if len(hole["bottoms"]) != len(layers):
            raise ValueError(f"{hole['id']}: one bottom elevation per layer is required.")
        if previous_x is not None and hole["x"] <= previous_x:
            raise ValueError("Borehole x coordinates must be distinct.")
        previous_x = hole["x"]
        interfaces = [hole["ground"], *hole["bottoms"]]
        # 标高向上为正，因此合法层序要求层顶不低于层底。
        # 使用 < 而不是 <= 是零厚度缺层能够保留的关键，不能改回严格正厚度。
        if any(top < bottom for top, bottom in zip(interfaces, interfaces[1:])):
            raise ValueError(f"{hole['id']}: 层底不能高于层顶（layer bottoms must descend or coincide）。")
        if interfaces[0] <= interfaces[-1]:
            raise ValueError(f"{hole['id']}: 整个钻孔的总厚度必须大于零。")
    ground = [(h["x"], h["ground"]) for h in holes]
    base = [(h["x"], h["bottoms"][-1]) for h in holes]
    domain = Polygon(ground + base[::-1])
    regions = []
    for index, layer in enumerate(layers):
        top = [(h["x"], h["ground"] if index == 0 else h["bottoms"][index - 1]) for h in holes]
        bottom = [(h["x"], h["bottoms"][index]) for h in holes]
        # Assemble interval-by-interval: a layer may pinch out and reappear.
        # A single whole-profile ring can self-touch when boundaries coincide.
        parts = []
        # 逐孔间区间建面，而非拼一个横贯全剖面的外环。
        # 尖灭点两条界线重合时，整剖面外环可能自接触；分段合并可正确表示再次出现的层。
        for i in range(len(holes) - 1):
            if top[i][1] == bottom[i][1] and top[i+1][1] == bottom[i+1][1]:
                continue
            part = Polygon([top[i], top[i+1], bottom[i+1], bottom[i]])
            if not part.is_valid or part.area <= 0:
                raise ValueError(f"Invalid geometry for layer {layer['id']} between boreholes {i+1} and {i+2}.")
            parts.append(part)
        polygon = unary_union(parts) if parts else Polygon()
        if not polygon.is_valid:
            raise ValueError(f"Invalid geometry for layer {layer['id']}.")
        regions.append({"layer": layer, "polygon": polygon})
    union = unary_union([r["polygon"] for r in regions])
    overlap = sum(regions[i]["polygon"].intersection(regions[j]["polygon"]).area
                  for i in range(len(regions)) for j in range(i))
    gap = domain.difference(union).area
    outside = union.difference(domain).area
    # 同时控制层间重叠、域内空洞和域外多余面积；容差随总面积缩放。
    # 几何面积一致只能证明分层覆盖正确，并不证明材料、水力边界或网格有效。
    tolerance = max(domain.area, 1.0) * 1e-10
    if max(overlap, gap, outside) > tolerance:
        raise ValueError("Geometry coverage validation failed.")
    report = {
        "status": "offline_geometry_validated", "units": "m",
        "interpolation": "piecewise linear; no extrapolation beyond end boreholes",
        "boreholes": len(holes), "regions": len(regions),
        "nonempty_layers": sum(not r['polygon'].is_empty for r in regions),
        "all_zero_layer_ids": [r['layer']['id'] for r in regions if r['polygon'].is_empty],
        "zero_thickness_at_boreholes": sum(
            top == bottom for h in holes
            for top, bottom in zip([h['ground'], *h['bottoms']], h['bottoms'])),
        "domain_area_m2": domain.area,
        "region_areas_m2": {r["layer"]["id"]: r["polygon"].area for r in regions},
        "overlap_area_m2": overlap, "gap_area_m2": gap, "outside_area_m2": outside,
        "area_tolerance_m2": tolerance,
        "plaxis_import_verified": False, "plaxis_mesh_verified": False,
        "note": "Demonstration geometry only. No constitutive, groundwater, or calculation parameters assigned."
    }
    return holes, domain, regions, report


# 把 Polygon/MultiPolygon 统一为可迭代的面列表；空层返回空列表。
# 调用者只对实际有面积的面出图或生成命令，逻辑层位仍在原 regions 中保留。
def polygon_parts(geometry):
    if geometry.is_empty:
        return []
    return [geometry] if geometry.geom_type == 'Polygon' else list(geometry.geoms)


# 输入排序钻孔和分层多边形，输出一张离线地层剖面图。
# 优先选本机中文字体，按层着色并标出孔位；多区域同层仅保留一个图例项。
def render_preview(holes, regions, output):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for font in ("Microsoft YaHei", "SimHei", "Noto Sans CJK SC"):
        if font in installed:
            plt.rcParams["font.sans-serif"] = [font]
            break
    plt.rcParams["axes.unicode_minus"] = False
    fig, ax = plt.subplots(figsize=(16, 7) if len(regions) > 10 else (12, 6), dpi=160)
    for region in regions:
        for index, part in enumerate(polygon_parts(region['polygon'])):
            x, y = part.exterior.xy
            ax.fill(x, y, color=region["layer"]["color"], edgecolor="#544A42", linewidth=0.6,
                    label=region["layer"]["name"] if index == 0 else None)
    for hole in holes:
        ax.plot([hole["x"], hole["x"]], [hole["bottoms"][-1], hole["ground"]],
                color="#355C70", linestyle="--", linewidth=1)
        ax.scatter([hole["x"]] * (len(hole["bottoms"]) + 1),
                   [hole["ground"], *hole["bottoms"]], color="#355C70", s=16, zorder=4)
        ax.annotate(f"{hole['id']} / x={hole['x']:g} m", (hole["x"], hole["ground"]),
                    xytext=(0, 10), textcoords="offset points", ha="center", fontsize=9)
    ax.set_aspect("equal", adjustable="box")
    ax.set_xlabel("剖面距离 x（m）")
    ax.set_ylabel("标高 y（m）")
    ax.set_title("钻孔驱动的二维土层 · 直线连接同名界面\n仅表示输入几何；未设置材料参数或水条件", pad=25)
    ax.margins(x=0.07, y=0.18)
    ax.grid(alpha=0.15)
    if len(regions) > 10:
        ax.legend(loc='upper left', bbox_to_anchor=(1.01, 1), fontsize=8)
    else:
        ax.legend(loc="lower right")
    fig.tight_layout()
    fig.savefig(output, bbox_inches="tight")
    plt.close(fig)


# 读取配置文件并生成几何验证、GeoJSON、命令文本及预览。
# 输出的 polygon 顺序通过 region_order.json 映射到逻辑层位；它不等于未来软件切分后的土簇顺序。
def generate(config, output):
    data = json.loads(Path(config).read_text(encoding="utf-8-sig"))
    holes, domain, regions, report = build_model(data)
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    # 将同一次离线生成的中间数据写成 UTF-8 JSON，保留中文，方便审阅和后续工作进程读取。
    def write_json(name, value):
        (output / name).write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    write_json("input_snapshot.json", data)
    write_json("validation.json", report)
    write_json("strata.geojson", {"type": "FeatureCollection", "features": [
        {"type": "Feature", "properties": {"id": r["layer"]["id"], "name": r["layer"]["name"],
                                            "area_m2": r["polygon"].area},
         "geometry": mapping(r["polygon"])} for r in regions]})
    # Bare PLAXIS commands, without comments or destructive project creation commands.
    commands = ["gotostructures"]
    xmin, ymin, xmax, ymax = domain.bounds
    commands.append(f"initializerectangular SoilContour {xmin:.15g} {ymin:.15g} {xmax:.15g} {ymax:.15g}")
    for r in regions:
        for part in polygon_parts(r['polygon']):
            coordinates = list(part.exterior.coords)[:-1]
            commands.append("polygon " + " ".join(f"({x:.15g} {y:.15g})" for x, y in coordinates))
    (output / "plaxis_commands.txt").write_text("\n".join(commands) + "\n", encoding="utf-8")
    exported_regions = [r for r in regions for part in polygon_parts(r['polygon'])]
    write_json("region_order.json", [{"command_polygon_order": i + 1, "layer_id": r["layer"]["id"],
                                     "layer_name": r["layer"]["name"]} for i, r in enumerate(exported_regions)])
    render_preview(holes, regions, output / "strata_preview.png")
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path(__file__).with_name("example_boreholes.json"))
    parser.add_argument("--output", type=Path, default=Path(__file__).with_name("output"))
    args = parser.parse_args()
    try:
        result = generate(args.config, args.output)
    except (ValueError, KeyError, TypeError) as error:
        parser.exit(2, f"Input/geometry error: {error}\n")
    print(json.dumps(result, ensure_ascii=False, indent=2))
