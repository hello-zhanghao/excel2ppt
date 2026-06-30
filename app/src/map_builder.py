from io import BytesIO
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import Normalize
import numpy as np
try:
    import cartopy.crs as ccrs
    import cartopy.feature as cfeature
    HAS_CARTOPY = True
except ImportError:
    HAS_CARTOPY = False

FONT_NAME = "Microsoft YaHei"
plt.rcParams["font.sans-serif"] = [FONT_NAME, "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

HW_DARK = "#182B49"
HW_RED = "#C8102E"


def build_scatter_map(df, lat_col, lon_col, metric_col, title="", color_hex=None):
    lat = df[lat_col].dropna().to_numpy()
    lon = df[lon_col].dropna().to_numpy()
    vals = df[metric_col].dropna().to_numpy()
    if len(lat) == 0:
        return None

    fig, ax = _create_map_ax(lat, lon)
    sizes = _normalize(vals, 60, 350)

    if color_hex:
        sc = ax.scatter(lon, lat, s=sizes, c=color_hex,
                        alpha=0.85, edgecolors="#333333", linewidth=1.0,
                        transform=ccrs.PlateCarree(), zorder=5)
    else:
        norm = Normalize(vmin=vals.min(), vmax=vals.max())
        sc = ax.scatter(lon, lat, s=sizes, c=vals, cmap="RdYlBu_r", norm=norm,
                        alpha=0.85, edgecolors="#333333", linewidth=1.0,
                        transform=ccrs.PlateCarree(), zorder=5)
        cbar = fig.colorbar(sc, ax=ax, shrink=0.78, pad=0.04,
                            orientation="vertical")
        cbar.set_label(metric_col, fontsize=11, color=HW_DARK, fontweight="bold")
        cbar.ax.tick_params(labelsize=9, colors=HW_DARK)

    _add_annotations(ax, df, lon, lat)

    if title:
        ax.set_title(title, fontsize=15, fontweight="bold", color=HW_DARK, pad=12)

    fig.tight_layout(pad=0.5)
    return fig


def build_heatmap(df, lat_col, lon_col, metric_col, title=""):
    lat = df[lat_col].dropna().to_numpy()
    lon = df[lon_col].dropna().to_numpy()
    vals = df[metric_col].dropna().to_numpy()
    if len(lat) == 0:
        return None

    fig, ax = _create_map_ax(lat, lon)
    sizes = _normalize(vals, 120, 550)
    alphas = np.clip(_normalize(vals, 0.08, 0.35), 0.02, 0.55)
    norm = Normalize(vmin=vals.min(), vmax=vals.max())

    sc = ax.scatter(lon, lat, s=sizes, c=vals, cmap="YlOrRd", norm=norm,
                    alpha=alphas, edgecolors="none",
                    transform=ccrs.PlateCarree(), zorder=5)
    cbar = fig.colorbar(sc, ax=ax, shrink=0.78, pad=0.04, orientation="vertical")
    cbar.set_label(metric_col, fontsize=11, color=HW_DARK, fontweight="bold")
    cbar.ax.tick_params(labelsize=9, colors=HW_DARK)

    if title:
        ax.set_title(title, fontsize=15, fontweight="bold", color=HW_DARK, pad=12)

    fig.tight_layout(pad=0.5)
    return fig


def save_map_image(fig, dpi=180):
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white", edgecolor="none",
                bbox_inches="tight", pad_inches=0.1)
    buf.seek(0)
    plt.close(fig)
    from PIL import Image
    img = Image.open(buf).convert("RGB")
    out = BytesIO()
    img.save(out, format="PNG")
    out.seek(0)
    return out


def _create_map_ax(lat, lon):
    margin = max((lon.max() - lon.min()) * 0.2, 0.01)
    lat_margin = max((lat.max() - lat.min()) * 0.2, 0.01)

    fig = plt.figure(figsize=(12, 7), facecolor="white")

    if HAS_CARTOPY:
        ax = fig.add_subplot(1, 1, 1, projection=ccrs.PlateCarree())
        ax.set_extent([lon.min() - margin, lon.max() + margin,
                       lat.min() - lat_margin, lat.max() + lat_margin],
                      crs=ccrs.PlateCarree())

        ax.add_feature(cfeature.LAND, facecolor="#F2F0EB", zorder=0)
        ax.add_feature(cfeature.OCEAN, facecolor="#DCE6F0", zorder=0)
        ax.add_feature(cfeature.LAKES, facecolor="#DCE6F0", zorder=0)
        ax.add_feature(cfeature.COASTLINE, linewidth=0.6, edgecolor="#888888", zorder=1)
        ax.add_feature(cfeature.BORDERS, linewidth=0.3, edgecolor="#BBBBBB", zorder=1)

        gl = ax.gridlines(draw_labels=True, linewidth=0.4, color="#CCCCCC",
                          alpha=0.6, linestyle="--")
        gl.top_labels = False
        gl.right_labels = False
        gl.xlabel_style = {"size": 8, "color": "#999999"}
        gl.ylabel_style = {"size": 8, "color": "#999999"}
    else:
        ax = fig.add_subplot(1, 1, 1)
        ax.set_xlim(lon.min() - margin, lon.max() + margin)
        ax.set_ylim(lat.min() - lat_margin, lat.max() + lat_margin)
        ax.set_xlabel("Longitude", fontsize=9, color="#999999")
        ax.set_ylabel("Latitude", fontsize=9, color="#999999")
        ax.tick_params(labelsize=8, colors="#999999")
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.grid(True, linestyle="--", alpha=0.3, color="#CCCCCC", linewidth=0.5)
        ax.set_facecolor("#F2F0EB")

    return fig, ax


def _add_annotations(ax, df, lon, lat):
    site_col = _find_col(df, ["site_id", "站点ID"])
    cell_col = _find_col(df, ["cell_id", "小区ID"])
    annotate = df[site_col].fillna("").tolist() if site_col else (
        df[cell_col].fillna("").tolist() if cell_col else []
    )
    for i in range(min(len(annotate), 6)):
        ax.annotate(
            str(annotate[i]), (lon[i], lat[i]),
            textcoords="offset points", xytext=(5, 5),
            fontsize=7, color=HW_DARK, fontweight="bold",
            bbox=dict(boxstyle="round,pad=0.15", facecolor="white",
                      alpha=0.75, edgecolor="none"),
            zorder=10,
            transform=ccrs.PlateCarree() if HAS_CARTOPY else ax.transData,
        )


def _normalize(values, lo, hi):
    v_min, v_max = values.min(), values.max()
    if v_max == v_min:
        return np.full(len(values), float(lo + hi) / 2)
    return lo + (values - v_min) / (v_max - v_min) * (hi - lo)


def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None
