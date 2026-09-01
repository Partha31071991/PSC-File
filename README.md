# Stock Statement AI Compiler — Secondary Compilation Fix

## Fixes included
- Corrected PDF SEC/Closing column mapping: Sales = 4th operational value and Closing = 5th.
- Browser-side secondary compilation now works locally from extracted PDF/Excel/CSV rows; `/api/analyze` is only a fallback for unusual layouts.
- Prevented matched rows from also being returned as review rows in the API.
- Retained SKU family/dose protection and the existing 85% matching rule.
- Retained the `/api` folder, Vercel configuration and master template for deployment compatibility.

## Deployment
Upload the **contents of this folder** (including `api/`, `vercel.json`, and `master-template.xlsx`) to Vercel.
