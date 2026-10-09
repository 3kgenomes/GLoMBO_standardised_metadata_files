# Standardised sampler CSV contract (draft)

Use [the header-only template](../schema/edna_sampler_standardised_TEMPLATE.csv)
as the proposed DAP-to-IDC exchange format. Its grain is **one row per sampling
attempt**, including failures, aborts and retries. `sample_event_id` identifies
an attempt; `filter_id` is the physical filter identifier; `filter_position` is
the sampler slot; `kit_uid` is the reconciled Wilderlab physical bag or kit
number. These are distinct identifiers.

The full field list, suggested source/owner and priority are in
[DAP_field_dictionary.csv](../schema/DAP_field_dictionary.csv). DAP can supply
instrument fields first; the processing team can add Kit/UID, vessel position
and curated terms before using the output scripts.

## Important formats and checks

| Fields | Required for | Format or rule |
| --- | --- | --- |
| `record_status`, `sample_event_id`, `filter_position`, `sample_start_utc`, `sample_end_utc`, `sampling_status`, `error_code` | CSV ingestion | Nonempty event ID, unique across attempts. Start/end ISO 8601 UTC end in `Z`. `sampling_status` is `complete`, `failed` or `aborted`; complete requires error code `0`, incomplete requires nonzero. |
| `sampled_volume_ml` | Both outputs | Number in mL when measured; empty if unknown. Do not substitute `target_volume_ml`. |
| `kit_uid` | Wilderlab and result reconciliation | Unique verified 5 or 6 digit physical UID for each submitted sample. Never derive from filter position. |
| `environment_type` | Wilderlab | Must match an option in the supplied Wilderlab workbook's `Lists` sheet. |
| `collection_date` | Wilderlab | `YYYY-MM-DD` calendar date, for example `2030-01-15`. Preserve it as text in the CSV; spreadsheet applications may reformat dates. |
| `latitude`, `longitude`, `coordinate_system` | Mapping and FAIRe | Paired WGS84 decimal degrees, latitude -90..90 and longitude -180..180. Leave the pair blank when unverified. |
| `geo_loc_name`, `env_broad_scale`, `env_local_scale`, `env_medium` | FAIRe completion | Confirmed geographic description and environment vocabulary terms; absence is reported for review. |
| `sample_category`, `neg_cont_type`, `pos_cont_type` | Controls | Supported categories: `sample`, `negative control`, `positive control`, `PCR standard`. Provide the relevant control type when applicable. |
| `sampling_depth_m` | Optional mapping | Depth in metres, 0..11000 when present. |

The CSV must have the complete header line. Most enrichment fields can remain
blank until verified, but the export scripts will reject completed samples
missing fields they need. Failed and aborted attempts stay in the CSV for
provenance but are excluded from the generated submission rows.

The scripts read UTF-8 CSV (with or without a UTF-8 BOM). Keep a distinct
`source_log_file`, vessel position timestamp and profile URI where these are
known. Do not manufacture missing coordinates or environmental ontology terms.

## Additional sources

The Wilderlab job JSON supplies contact, country, assay panel, sharing choice
and billing details where required. The FAIRe project JSON supplies confirmed
project metadata. Neither is an instrument log. Later Wilderlab result
`metadata.UID` and `metadata.ClientSampleID` must reconcile to `kit_uid` and
`sample_event_id` respectively.

The initial FAIRe output is a draft. Assay codes and taxonomy come from the
later results stage; extraction, run, database and sequence file provenance
must be obtained and reviewed separately for a complete publication record.
