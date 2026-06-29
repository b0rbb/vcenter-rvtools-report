#!/usr/bin/env python3
"""
generate_vcenter_report.py
--------------------------
Reads an RVTools XLSX export and produces a PowerPoint report with:
  - A divider slide per Datacenter
  - ESXi Hosts slide: one row per host with CPU model, memory, VM count, etc.
  - VM Resource Summary slide: aggregated storage allocated/used, memory, vCPUs, and p:v CPU ratio

Usage:
    python generate_vcenter_report.py <input.xlsx> [output.pptx]

Requires:
    pip install openpyxl python-pptx
"""

import sys
import math
from pathlib import Path
from collections import defaultdict

import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
DARK_BG   = RGBColor(0x1E, 0x29, 0x3B)   # slide background / divider fill
ACCENT    = RGBColor(0x00, 0x8C, 0xD7)   # VMware-ish blue
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY= RGBColor(0xF4, 0xF6, 0xF8)
MID_GRAY  = RGBColor(0xBD, 0xC3, 0xC7)
DARK_TEXT = RGBColor(0x1A, 0x1A, 0x2E)
ROW_ALT   = RGBColor(0xEA, 0xF4, 0xFB)   # alternating row tint

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------
def mib_to_gib(mib):
    if mib is None: return 0.0
    return round(float(mib) / 1024, 1)

def mib_to_tib(mib):
    if mib is None: return 0.0
    return round(float(mib) / (1024 * 1024), 2)

def mb_to_gib(mb):
    """vHostMemorySize is in MB."""
    if mb is None: return 0.0
    return round(float(mb) / 1024, 1)

def shorten_host(name):
    """Return the short hostname (part before first '.')."""
    return name.split(".")[0] if name else name

def ratio_str(phys, virt):
    if not phys: return "N/A"
    r = virt / phys
    return f"1 : {r:.1f}"


# ---------------------------------------------------------------------------
# Column-name aliases  (old RVTools camelCase  →  new human-readable)
# ---------------------------------------------------------------------------
HOST_COL_ALIASES = {
    # canonical_key: [preferred_name, fallback_name, ...]
    "vHostName":        ["vHostName",        "Host"],
    "vHostDatacenter":  ["vHostDatacenter",  "Datacenter"],
    "vHostCluster":     ["vHostCluster",     "Cluster"],
    "vHostCpuModel":    ["vHostCpuModel",    "CPU Model"],
    "vHostCpuMhz":      ["vHostCpuMhz",      "Speed"],
    "vHostNumCpu":      ["vHostNumCpu",      "# CPU"],
    "vHostCoresPerCPU": ["vHostCoresPerCPU", "Cores per CPU"],
    "vHostNumCpuCores": ["vHostNumCpuCores", "# Cores"],
    "vHostMemorySize":  ["vHostMemorySize",  "# Memory"],
    "vHostVMsTotal":    ["vHostVMsTotal",    "# VMs total"],
    "vHostVMs":         ["vHostVMs",         "# VMs"],
    "vHostvCPUs":       ["vHostvCPUs",       "# vCPUs"],
    "vHostvRAM":        ["vHostvRAM",        "vRAM"],
}

INFO_COL_ALIASES = {
    "vInfoVMName":              ["vInfoVMName",              "VM"],
    "vInfoTemplate":            ["vInfoTemplate",            "Template"],
    "vInfoCPUs":                ["vInfoCPUs",                "CPUs"],
    "vInfoMemory":              ["vInfoMemory",              "Memory"],
    "vInfoTotalDiskCapacityMiB":["vInfoTotalDiskCapacityMiB","Total disk capacity MiB"],
    "vInfoProvisioned":         ["vInfoProvisioned",         "Provisioned MiB"],
    "vInfoInUse":               ["vInfoInUse",               "In Use MiB"],
    "vInfoDataCenter":          ["vInfoDataCenter",          "Datacenter"],
    "vInfoCluster":             ["vInfoCluster",             "Cluster"],
    "vInfoHost":                ["vInfoHost",                "Host"],
}

def resolve_index(header_map, aliases, canonical):
    """Return the column index for the first alias found in header_map."""
    for name in aliases[canonical]:
        if name in header_map:
            return header_map[name]
    raise KeyError(f"Could not find column for '{canonical}'. "
                   f"Tried: {aliases[canonical]}")

def make_getter(header_map, aliases):
    """Return a callable g(row, canonical_key) → value."""
    idx_cache = {}
    def g(row, key):
        if key not in idx_cache:
            idx_cache[key] = resolve_index(header_map, aliases, key)
        v = row[idx_cache[key]]
        return v
    return g


# ---------------------------------------------------------------------------
# Read data from XLSX
# ---------------------------------------------------------------------------
def load_workbook_data(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=False, data_only=True)

    # ---- vHost ----
    ws_host = wb["vHost"]
    host_rows = list(ws_host.iter_rows(values_only=True))
    hi = {h: i for i, h in enumerate(host_rows[0])}
    hget = make_getter(hi, HOST_COL_ALIASES)

    hosts = []
    for row in host_rows[1:]:
        if not hget(row, "vHostName"):
            continue
        hosts.append({
            "name":             hget(row, "vHostName"),
            "datacenter":       hget(row, "vHostDatacenter"),
            "cluster":          hget(row, "vHostCluster"),
            "cpu_model":        hget(row, "vHostCpuModel") or "",
            "cpu_mhz":          hget(row, "vHostCpuMhz") or 0,
            "num_sockets":      hget(row, "vHostNumCpu") or 0,
            "cores_per_socket": hget(row, "vHostCoresPerCPU") or 0,
            "total_cores":      hget(row, "vHostNumCpuCores") or 0,
            "memory_mb":        hget(row, "vHostMemorySize") or 0,
            "vms_total":        hget(row, "vHostVMsTotal") or 0,
            "vms_on":           hget(row, "vHostVMs") or 0,
            "vcpus":            hget(row, "vHostvCPUs") or 0,
            "vram_mib":         hget(row, "vHostvRAM") or 0,
        })

    # ---- vInfo ----
    ws_info = wb["vInfo"]
    info_rows = list(ws_info.iter_rows(values_only=True))
    ii = {h: i for i, h in enumerate(info_rows[0])}
    iget = make_getter(ii, INFO_COL_ALIASES)

    vms = []
    for row in info_rows[1:]:
        if not iget(row, "vInfoVMName"):
            continue
        # skip templates
        if str(iget(row, "vInfoTemplate")).lower() == "true":
            continue
        vms.append({
            "name":       iget(row, "vInfoVMName"),
            "datacenter": iget(row, "vInfoDataCenter"),
            "cluster":    iget(row, "vInfoCluster"),
            "host":       iget(row, "vInfoHost"),
            "cpus":       iget(row, "vInfoCPUs") or 0,
            "memory_mib": iget(row, "vInfoMemory") or 0,
            "disk_cap_mib":   iget(row, "vInfoTotalDiskCapacityMiB") or 0,
            "provisioned_mib":iget(row, "vInfoProvisioned") or 0,
            "inuse_mib":      iget(row, "vInfoInUse") or 0,
        })

    wb.close()
    return hosts, vms


# ---------------------------------------------------------------------------
# Aggregate per-datacenter
# ---------------------------------------------------------------------------
def aggregate(hosts, vms):
    dcs = sorted(set(h["datacenter"] for h in hosts if h["datacenter"]))

    dc_hosts = defaultdict(list)
    for h in hosts:
        dc_hosts[h["datacenter"]].append(h)

    dc_vms = defaultdict(list)
    for v in vms:
        dc_vms[v["datacenter"]].append(v)

    dc_summary = {}
    for dc in dcs:
        dc_vm_list = dc_vms[dc]
        dc_host_list = dc_hosts[dc]

        total_vcpus     = sum(v["cpus"] for v in dc_vm_list)
        total_vmem_mib  = sum(v["memory_mib"] for v in dc_vm_list)
        total_alloc_mib = sum(v["provisioned_mib"] for v in dc_vm_list)
        total_used_mib  = sum(v["inuse_mib"] for v in dc_vm_list)
        total_pcores    = sum(h["total_cores"] for h in dc_host_list)

        dc_summary[dc] = {
            "hosts":          dc_host_list,
            "vm_count":       len(dc_vm_list),
            "total_vcpus":    total_vcpus,
            "total_vmem_mib": total_vmem_mib,
            "alloc_mib":      total_alloc_mib,
            "used_mib":       total_used_mib,
            "total_pcores":   total_pcores,
        }

    return dcs, dc_summary


# ---------------------------------------------------------------------------
# PowerPoint helpers
# ---------------------------------------------------------------------------
def set_slide_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=WHITE,
                 align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb


def add_header_bar(slide, title, subtitle=None):
    """Dark top bar with title + optional subtitle."""
    bar = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(0), Inches(0), SLIDE_W, Inches(1.1)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = DARK_BG
    bar.line.fill.background()

    add_text_box(slide, title,
                 Inches(0.35), Inches(0.08), Inches(10), Inches(0.65),
                 font_size=26, bold=True, color=WHITE)
    if subtitle:
        add_text_box(slide, subtitle,
                     Inches(0.35), Inches(0.68), Inches(12), Inches(0.38),
                     font_size=13, bold=False, color=MID_GRAY)


def add_accent_line(slide, top_inches):
    line = slide.shapes.add_shape(1,
        Inches(0), Inches(top_inches), SLIDE_W, Pt(3))
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()


def cell_set(cell, text, font_size=10, bold=False,
             bg_color=None, font_color=DARK_TEXT,
             align=PP_ALIGN.LEFT):
    cell.text = str(text) if text is not None else ""
    tf = cell.text_frame
    tf.word_wrap = False
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = font_color
    if bg_color:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------
def add_divider_slide(prs, dc_name):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, DARK_BG)

    # Accent stripe on left
    stripe = slide.shapes.add_shape(1,
        Inches(0), Inches(0), Inches(0.18), SLIDE_H)
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = ACCENT
    stripe.line.fill.background()

    add_text_box(slide, dc_name.upper(),
                 Inches(0.5), Inches(2.8), Inches(12), Inches(1.5),
                 font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_text_box(slide, "Datacenter Report",
                 Inches(0.5), Inches(4.1), Inches(12), Inches(0.6),
                 font_size=20, bold=False, color=ACCENT, align=PP_ALIGN.CENTER)
    return slide


def add_hosts_slide(prs, dc_name, host_list):
    """One slide showing all ESXi hosts in this DC with key specs."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, LIGHT_GRAY)
    add_header_bar(slide, f"{dc_name}  —  ESXi Hosts",
                   f"{len(host_list)} host(s)")
    add_accent_line(slide, 1.1)

    # Column definitions: (header, width_inches, attr_fn)
    COL_DEFS = [
        ("Host",            2.5,  lambda h: shorten_host(h["name"])),
        ("Cluster",         1.8,  lambda h: h["cluster"] or ""),
        ("CPU Model",       3.2,  lambda h: h["cpu_model"]),
        ("Sockets × Cores", 1.4,  lambda h: f"{h['num_sockets']} × {h['cores_per_socket']}"),
        ("Total Cores",     1.0,  lambda h: h["total_cores"]),
        ("Freq (GHz)",      0.85, lambda h: f"{h['cpu_mhz']/1000:.2f}"),
        ("Memory (GB)",     1.0,  lambda h: mb_to_gib(h["memory_mb"])),
        ("VMs (total/on)",  1.1,  lambda h: f"{h['vms_total']} / {h['vms_on']}"),
        ("vCPUs",           0.75, lambda h: h["vcpus"]),
    ]

    n_rows = len(host_list) + 1  # +1 header
    n_cols = len(COL_DEFS)

    table_top    = Inches(1.25)
    table_left   = Inches(0.25)
    table_width  = SLIDE_W - Inches(0.5)
    row_height   = Inches(0.38)
    table_height = row_height * n_rows

    tbl = slide.shapes.add_table(
        n_rows, n_cols,
        table_left, table_top, table_width, table_height
    ).table

    # Set column widths
    total_w = sum(w for _, w, _ in COL_DEFS)
    scale = (table_width / Inches(1)) / total_w
    for ci, (_, w, _) in enumerate(COL_DEFS):
        tbl.columns[ci].width = Inches(w * scale)

    # Header row
    for ci, (hdr, _, _) in enumerate(COL_DEFS):
        cell_set(tbl.cell(0, ci), hdr,
                 font_size=10, bold=True,
                 bg_color=DARK_BG, font_color=WHITE,
                 align=PP_ALIGN.CENTER)

    # Data rows
    for ri, host in enumerate(host_list, start=1):
        bg = WHITE if ri % 2 == 1 else ROW_ALT
        for ci, (_, _, fn) in enumerate(COL_DEFS):
            val = fn(host)
            align = PP_ALIGN.LEFT if ci <= 2 else PP_ALIGN.CENTER
            cell_set(tbl.cell(ri, ci), val,
                     font_size=9, bold=False,
                     bg_color=bg, font_color=DARK_TEXT,
                     align=align)

    # Totals row
    # Build a "Total" summary line appended below the table as text
    total_cores = sum(h["total_cores"] for h in host_list)
    total_mem   = sum(mb_to_gib(h["memory_mb"]) for h in host_list)
    total_vms   = sum(h["vms_total"] for h in host_list)
    total_vcpus = sum(h["vcpus"] for h in host_list)

    summary = (f"Totals  —  Physical Cores: {total_cores}   "
               f"Memory: {total_mem:.1f} GB   "
               f"VMs (total): {total_vms}   "
               f"vCPUs assigned: {total_vcpus}")

    footer_top = table_top + table_height + Inches(0.12)
    add_text_box(slide, summary,
                 Inches(0.25), footer_top, SLIDE_W - Inches(0.5), Inches(0.35),
                 font_size=10, bold=True, color=DARK_TEXT, align=PP_ALIGN.LEFT)

    return slide


def add_summary_slide(prs, dc_name, summary):
    """Per-DC VM resource summary slide."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, LIGHT_GRAY)
    add_header_bar(slide, f"{dc_name}  —  VM Resource Summary",
                   f"{summary['vm_count']} virtual machine(s) (non-template)")
    add_accent_line(slide, 1.1)

    # Compute display values
    alloc_tib = mib_to_tib(summary["alloc_mib"])
    used_tib  = mib_to_tib(summary["used_mib"])
    vmem_gib  = mib_to_gib(summary["total_vmem_mib"])
    vcpus     = summary["total_vcpus"]
    pcores    = summary["total_pcores"]
    p2v       = ratio_str(pcores, vcpus)
    used_pct  = (summary["used_mib"] / summary["alloc_mib"] * 100
                 if summary["alloc_mib"] else 0)

    cards = [
        ("Storage Allocated (All VMs)",  f"{alloc_tib} TiB",
         f"({mib_to_gib(summary['alloc_mib']):.1f} GiB)"),
        ("Storage Used (All VMs)",        f"{used_tib} TiB",
         f"{used_pct:.1f}% of allocated"),
        ("Memory Allocated (All VMs)",    f"{vmem_gib:.1f} GiB",
         f"({summary['total_vmem_mib']:,} MiB)"),
        ("vCPUs Assigned (All VMs)",      f"{vcpus:,}",
         f"across {summary['vm_count']} VMs"),
        ("Physical CPU Cores (All Hosts)",f"{pcores:,}",
         f"across {len(summary['hosts'])} host(s)"),
        ("pCPU : vCPU Ratio",             p2v,
         "physical cores to virtual CPUs"),
    ]

    # 2-column card layout
    card_w   = Inches(5.9)
    card_h   = Inches(1.55)
    col_gap  = Inches(0.4)
    row_gap  = Inches(0.22)
    start_x  = [Inches(0.35), Inches(0.35) + card_w + col_gap]
    start_y  = Inches(1.3)

    for idx, (label, value, sub) in enumerate(cards):
        col = idx % 2
        row = idx // 2
        x = start_x[col]
        y = start_y + row * (card_h + row_gap)

        # Card background
        card = slide.shapes.add_shape(1, x, y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = ACCENT
        card.line.width = Pt(1.5)

        # Accent left bar on card
        bar = slide.shapes.add_shape(1, x, y, Inches(0.08), card_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = ACCENT
        bar.line.fill.background()

        # Label
        add_text_box(slide, label,
                     x + Inches(0.18), y + Inches(0.1),
                     card_w - Inches(0.25), Inches(0.38),
                     font_size=11, bold=False, color=MID_GRAY)

        # Value (big)
        add_text_box(slide, value,
                     x + Inches(0.18), y + Inches(0.42),
                     card_w - Inches(0.25), Inches(0.72),
                     font_size=28, bold=True, color=DARK_TEXT)

        # Sub-label
        add_text_box(slide, sub,
                     x + Inches(0.18), y + Inches(1.1),
                     card_w - Inches(0.25), Inches(0.38),
                     font_size=10, bold=False, color=MID_GRAY, italic=True)

    return slide


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def build_presentation(xlsx_path, pptx_path):
    print(f"Reading: {xlsx_path}")
    hosts, vms = load_workbook_data(xlsx_path)
    dcs, dc_summary = aggregate(hosts, vms)

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    # Cover slide
    blank = prs.slide_layouts[6]
    cover = prs.slides.add_slide(blank)
    set_slide_background(cover, DARK_BG)

    stripe = cover.shapes.add_shape(1,
        Inches(0), Inches(3.2), SLIDE_W, Inches(0.06))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = ACCENT
    stripe.line.fill.background()

    add_text_box(cover, "vCenter Infrastructure Report",
                 Inches(0.5), Inches(1.8), Inches(12), Inches(1.1),
                 font_size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_box(cover, f"{len(dcs)} Datacenter(s)  ·  {len(hosts)} ESXi Host(s)  ·  {len(vms)} VM(s)",
                 Inches(0.5), Inches(3.4), Inches(12), Inches(0.5),
                 font_size=16, bold=False, color=MID_GRAY, align=PP_ALIGN.CENTER)
    add_text_box(cover, "Generated from RVTools export",
                 Inches(0.5), Inches(6.8), Inches(12), Inches(0.4),
                 font_size=11, bold=False, color=MID_GRAY,
                 align=PP_ALIGN.CENTER, italic=True)

    # Per-datacenter slides grouped together
    for dc in dcs:
        s = dc_summary[dc]
        add_divider_slide(prs, dc)
        add_hosts_slide(prs, dc, s["hosts"])
        add_summary_slide(prs, dc, s)

    prs.save(pptx_path)
    print(f"Saved:   {pptx_path}")
    print(f"Slides:  {len(prs.slides)}  ({len(dcs)} DCs × 3 slides + 1 cover)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_vcenter_report.py <input.xlsx> [output.pptx]")
        sys.exit(1)

    xlsx = Path(sys.argv[1])
    pptx = Path(sys.argv[2]) if len(sys.argv) > 2 else xlsx.with_suffix(".pptx")
    build_presentation(str(xlsx), str(pptx))
