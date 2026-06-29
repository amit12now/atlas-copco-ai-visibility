# New Client Onboarding Checklist

## 1. Collect raw exports
Same 8 sources as Manchester, fresh for the new client:

- Semrush > Position Tracking > AI Overview filter → CSV
- Semrush > AI Visibility > Brand Topics → CSV
- Semrush > AI Visibility > Gap Topics → CSV
- Semrush > AI Visibility > Prompts detail → CSV
- Semrush > AI Visibility > Sources → CSV
- ChatGPT query-checker output → CSV
- GA4 Acquisition report (AI referral sources: chatgpt.com, gemini.google.com, perplexity.ai) → XLSX, "Sheet1"
- Search Console Performance, Pages/Countries/Devices tabs → one XLSX with 3 sheets

**Gotcha:** `create_new_template.py` has these exact filenames hardcoded (not patterns):
`Ai Overviews data.csv`, `brand_topics_manchester_ac_ae_ae.csv`, `gap_topics_manchester_ac_ae_ae.csv`,
`prompts_manchester_ac_ae_ae.csv`, `sources_manchester_ac_ae_ae.csv`, `chatgpt_results_uae.csv`,
`Analytics data.xlsx`, `search-console-data.xlsx`.
Either rename the new client's export files to match these exactly, or edit those filename strings near the top of `build()` in `create_new_template.py`.

## 2. New project folder + repo
- Duplicate the AI Project folder for this client.
- Remove the old client's raw CSVs/xlsx.
- Create a new **private** GitHub repo (per README's data-sensitivity rule — never reuse one client's repo for another).

## 3. Build the template
```
python create_new_template.py --source-dir <folder with new exports> --brand-domain <newclient-domain> --report-month YYYY-MM
python update_template.py "AI_Visibility_Tracking_Template.xlsx"
```
First command builds Settings/README/Mapping_Reference/RAW_*/empty Clean_* sheets.
Second populates Clean_* from RAW_*.

## 4. Re-brand dashboard.py
- Color constants (`PURPLE`, `YELLOW`, `GREY`, etc.) and `DONUT_PALETTE` → new client's brand colors.
- `.streamlit/config.toml` `[theme]` block → same colors.
- `RELEVANT_TOPIC_KEYWORDS` → new allowlist for the client's industry (currently education/MBA-specific).
- `TOPIC_DISPLAY` → clear it; add new renames only as needed.
- Double-check `Settings!Brand_Display_Name` reads correctly in the header/verdict text.

## 5. Test & deploy
- `streamlit run dashboard.py` locally — click every tab, sanity-check numbers against the raw exports.
- Push to the new private repo, deploy as its own separate Streamlit Cloud app.

## 6. Monthly cadence going forward
- Re-export raw data → paste into RAW_* tabs → bump `Settings!Report_Month` → `python update_template.py ...` → refresh dashboard.

---

## Starter prompt for next session
```
I'm onboarding a new client into the AI visibility dashboard template.
Client: <name>
Domain: <domain>
Industry: <industry>
Brand colors: <hex codes>
Raw exports are in: <folder path>
Walk me through building their template and re-branding the dashboard.
```
