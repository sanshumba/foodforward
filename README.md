# FoodForward Mother and Child Programme Dashboard

This folder contains a Python dashboard and Supabase upload script for the cleaned and anonymised Project 1 dataset.

## Files

- `app.py`  
  Streamlit dashboard for exploring the cleaned anonymised data.

- `app_dashboard2.py`  
  Separate Project 2 dashboard focused on basket-composition findings, site comparison, gaps, external context, and illustrative cost scenarios.

- `load_to_supabase.py`  
  Python script that uploads the cleaned Excel sheets into Supabase tables.

- `supabase_schema.sql`  
  SQL file for creating the required Supabase tables.

- `.env.example`  
  Example environment file showing the required Supabase credentials.

- `requirements.txt`  
  Python packages required to run the dashboard and upload script.

- `data/foodforward_project1_cleaned_anonymised_dataset.xlsx`  
  The cleaned anonymised Excel workbook created in Project 1.

## 1. Install requirements

From this folder, run:

```bash
pip install -r requirements.txt
```

## 2. Run the dashboard using the local Excel file

```bash
streamlit run app.py
```

By default, the dashboard reads:

```text
data/foodforward_project1_cleaned_anonymised_dataset.xlsx
```

## 3. Set up Supabase

Create a Supabase project, then open the SQL editor and run:

```sql
-- paste the contents of supabase_schema.sql here
```

This creates the following tables:

- `ff_sites`
- `ff_bo_profile`
- `ff_food_distributed`
- `ff_participants`
- `ff_measurements`
- `ff_data_quality_log`

## 4. Add Supabase credentials

Create a file called `.env` in this folder using `.env.example` as the template:

```text
SUPABASE_URL=https://your-project-id.supabase.co
SUPABASE_ANON_KEY=your_anon_key_for_dashboard_reads
SUPABASE_SERVICE_ROLE_KEY=your_service_role_key_for_uploads
```

Use the service role key only on your own computer or trusted server. Do not expose it publicly.

## 5. Upload the cleaned data to Supabase

```bash
python load_to_supabase.py
```

The script reads the cleaned Excel workbook, normalises the column names to snake_case, and inserts each sheet into the matching Supabase table.

## 6. Run the dashboard from Supabase

```bash
streamlit run app.py
```

In the dashboard sidebar, choose `Supabase` as the data source.

## Dashboard sections

The dashboard includes:

1. **Overview**
   - Number of sites
   - Number of participants
   - Distribution records
   - Total distributed line weight
   - Measurement records
   - Participants by site
   - Data quality issues

2. **Distribution**
   - Total quantity
   - Total weight
   - Unique items
   - Distribution by month
   - Distribution by category
   - Top distributed items

3. **Measurements**
   - Participants measured
   - Average weight
   - Average height
   - Flagged records
   - Average weight over time
   - Average height over time
   - Weight-height relationship

4. **Site profile**
   - BO/site profile summary
   - Total beneficiaries by site
   - BO size distribution

5. **Data quality**
   - Logged issues
   - Participants affected
   - Issue frequencies
   - Full data quality log

## Privacy note

The dashboard uses the cleaned anonymised Project 1 dataset. Direct child/client names, direct lookup values, exact dates of birth, and GPS coordinates are not displayed.
