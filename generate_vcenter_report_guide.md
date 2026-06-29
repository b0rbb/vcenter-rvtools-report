# vCenter PowerPoint Report Generator

Generates a PowerPoint slide deck from an RVTools XLSX export, with slides grouped by datacenter covering ESXi host specs and VM resource summaries.

## Requirements

- Python 3.7+
- Two Python libraries:

```bash
pip install openpyxl python-pptx
```

## Usage

```bash
python generate_vcenter_report.py <input.xlsx> [output.pptx]
```

- `<input.xlsx>` — required. Path to your RVTools XLSX export.
- `[output.pptx]` — optional. Output file path. If omitted, the output is saved in the same location as the input file with a `.pptx` extension.

**Examples:**

```bash
# Output saved as vCenter_export.pptx
python generate_vcenter_report.py vCenter_export.xlsx

# Output saved to a specific path
python generate_vcenter_report.py vCenter_export.xlsx ~/Desktop/report.pptx
```

## What the script produces

The deck contains one set of 3 slides per datacenter, grouped together:

1. **Divider slide** — datacenter name
2. **ESXi Hosts** — table with one row per host showing CPU model, socket/core counts, memory, VM count, and vCPUs assigned
3. **VM Resource Summary** — six metric cards: storage allocated, storage used, memory allocated, vCPUs assigned, physical CPU cores, and the pCPU:vCPU ratio

## Exporting from RVTools

1. Open RVTools and connect to your vCenter server
2. Go to **File → Export all to xlsx**
3. Save the file and pass it to this script

The script works with both the older RVTools camelCase column format and the newer human-readable column format, so any recent version of RVTools should work.

## Notes

- Templates are automatically excluded from VM counts and resource totals
- The script loads the workbook with filtered/hidden rows visible, so all hosts are captured even if Excel filters are active on the sheet
- Datacenters are sorted alphabetically in the deck
