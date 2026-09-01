FINAL VERIFICATION REPORT

Identity logic:
1) First meaningful statement headline is primary.
2) Filename is fallback when it contains a meaningful stockist name.
3) Detected stockist is matched against CUSTOMER NAME in master.
4) HQ is mapped from the matched master row even when HQ is absent from the statement.
5) Weak/non-master names are never silently mapped.

SKU business rules:
- 85% is the default business threshold.
- >=85% accepted automatically.
- <85% remains review/ignore when unresolved.
- GLUCORYL MV, MD and MP are separate SKU families during auto-matching.
- Forte/non-Forte and dose are protected for GlucoryL M/MV.
- Exact authoritative Glucoryl M list: 8 SKUs.
- Exact authoritative Glucoryl MV list: 5 SKUs.
- Repeated source SKU rows are not combined during review/output canonicalization.

Product group rules:
- Glucoryl M: exactly the 8 supplied SKUs.
- Glucoryl MV: exactly the 5 supplied SKUs.
- MD/MP are not Glucoryl M.
- MV cannot fall into M.
- Portal and downloaded Excel use the same group order/keys.

Analysis:
- Output, Stockist Analysis, HQ Analysis, Product Analysis and SKU Group Analysis use canonical matched rows.
- Inventory days use 30-day basis.
- >40 days is critical.

Formats:
PDF, XLSX, XLS, CSV, TXT, HTML/HTM.

Technical checks:
- Python API syntax: PASS
- index.html JavaScript syntax: PASS
- API accepts compact JSON identify/analyze requests.
- Existing multipart API compatibility retained.
- Mobile responsive CSS present.
