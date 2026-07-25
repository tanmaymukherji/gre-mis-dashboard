import argparse
import csv
import hashlib
import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd


TRADER_ID = "738"
ORG_NAME = "GRE Curated Financing Solutions"
DEFAULT_VALUECHAIN = "Financial"
DEFAULT_APPLICATION = "Loans and Schemes"
DEFAULT_CATEGORY = "Service offerings"
DEFAULT_GROUP = "Service"
DEFAULT_TYPE = "Financial Support"
DEFAULT_6M = "Money"
DEFAULT_LANGUAGE = "English"

STATE_ALIASES = {
    "all": "India",
    "india": "India",
    "orissa": "Odisha",
    "odisha": "Odisha",
    "pondicherry": "Puducherry",
    "delhi": "Delhi",
}

KNOWN_INDIAN_STATES = {
    "Andaman and Nicobar Islands",
    "Andhra Pradesh",
    "Arunachal Pradesh",
    "Assam",
    "Bihar",
    "Chandigarh",
    "Chhattisgarh",
    "Dadra and Nagar Haveli and Daman and Diu",
    "Delhi",
    "Goa",
    "Gujarat",
    "Haryana",
    "Himachal Pradesh",
    "Jammu and Kashmir",
    "Jharkhand",
    "Karnataka",
    "Kerala",
    "Ladakh",
    "Lakshadweep",
    "Madhya Pradesh",
    "Maharashtra",
    "Manipur",
    "Meghalaya",
    "Mizoram",
    "Nagaland",
    "Odisha",
    "Puducherry",
    "Punjab",
    "Rajasthan",
    "Sikkim",
    "Tamil Nadu",
    "Telangana",
    "Tripura",
    "Uttar Pradesh",
    "Uttarakhand",
    "West Bengal",
}


IMPORT_COLUMNS = [
    "SourceSchemeId",
    "SourceSlug",
    "TraderId",
    "OrganisationName",
    "SolutionName",
    "OfferingName",
    "OfferingCategory",
    "OfferingGroup",
    "OfferingType",
    "6M",
    "PrimaryValuechain",
    "PrimaryApplication",
    "Valuechains",
    "Applications",
    "Tags",
    "Languages",
    "Geographies",
    "AboutSolution",
    "AboutOffering",
    "Who Can avail it",
    "Trainer Name",
    "Trainer Email Address",
    "Trainer Phone Number",
    "Trainer Details",
    "Duration",
    "Prerequisites - Participants and Training",
    "Cost (Service)",
    "Support post Service",
    "Support post Service Cost",
    "Is it offered - Online or Offline",
    "Certification Offered",
    "Remarks on Cost",
    "Location Availability",
    "Service offering Brochure",
    "Contact Details",
    "Offering Link on GRE",
    "SourceSection",
    "SourceProvider",
    "SourceGovernmentLink",
    "SourceSchemeInfoLink",
    "SourceEnrollmentLink",
    "ImportNotes",
]


def clean(value):
    if value is None or pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("_x000D_", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_url(value):
    text = clean(value)
    if not text:
        return ""
    parts = [part.strip() for part in re.split(r"\s*\|\s*|,\s+(?=https?://)", text) if part.strip()]
    for part in parts:
        if re.match(r"^https?://", part, flags=re.I):
            return part
    if re.match(r"^www\.", text, flags=re.I):
        return f"https://{text}"
    if re.match(r"^https?://", text, flags=re.I):
        return text
    return text


def split_terms(value):
    text = clean(value)
    if not text:
        return []
    text = text.replace("#", "|")
    text = re.sub(r"\s+\|\s+", "|", text)
    parts = re.split(r"[;,\n|]+", text)
    return unique([part.strip(" .#-") for part in parts if part.strip(" .#-")])


def unique(values):
    seen = set()
    out = []
    for value in values:
        text = clean(value)
        key = text.lower()
        if text and key not in seen:
            out.append(text)
            seen.add(key)
    return out


def normalize_geographies(row):
    states = clean(row.get("States"))
    resolved = clean(row.get("Scheme States (resolved)"))
    if states.lower() == "all" or not resolved:
        return ["India"], []
    pieces = re.split(r"[;,|/]+|\band\b", resolved)
    geographies = []
    warnings = []
    for piece in pieces:
        normalized = clean(piece)
        if not normalized:
            continue
        normalized = STATE_ALIASES.get(normalized.lower(), normalized)
        geographies.append(normalized)
        if normalized != "India" and normalized not in KNOWN_INDIAN_STATES:
            warnings.append(f"Unknown geography: {normalized}")
    return unique(geographies or ["India"]), unique(warnings)


def infer_delivery_mode(row):
    enrollment = clean(row.get("Link for enrollment"))
    govt = clean(row.get("Link to government website"))
    info = clean(row.get("Link to scheme information"))
    if re.search(r"https?://|www\\.", " ".join([enrollment, govt, info]), flags=re.I):
        return "Online"
    if enrollment:
        return "Offline"
    return "Online / Offline"


def build_description(row):
    sections = [
        ("Summary", row.get("Short Description")),
        ("Details", row.get("Details")),
        ("Introduction", row.get("Introduction")),
        ("Benefits", row.get("Benefits")),
        ("Eligibility", row.get("Eligibility")),
        ("Target audience", row.get("Target Audience")),
        ("Impact on wMSMEs", row.get("Impact on wMSMEs")),
        ("Provider", row.get("Scheme Provider")),
        ("Implementing agency", row.get("Implementing Agency")),
        ("Ministry", row.get("Name of Ministry")),
        ("Enrollment link", clean_url(row.get("Link for enrollment"))),
        ("Scheme information", clean_url(row.get("Link to scheme information"))),
        ("Government website", clean_url(row.get("Link to government website"))),
        ("WEP resource", clean_url(row.get("Scheme URL (wep.gov.in resource-detail)"))),
    ]
    lines = []
    for label, value in sections:
        text = clean(value)
        if text:
            lines.append(f"{label}: {text}")
    return "\n\n".join(lines)


def build_trainer_details(row):
    parts = [
        ("Scheme provider", row.get("Scheme Provider")),
        ("Implementing agency", row.get("Implementing Agency")),
        ("Ministry", row.get("Name of Ministry")),
        ("Enrollment", clean_url(row.get("Link for enrollment"))),
        ("Scheme information", clean_url(row.get("Link to scheme information"))),
        ("Government website", clean_url(row.get("Link to government website"))),
    ]
    return "\n".join(f"{label}: {clean(value)}" for label, value in parts if clean(value))


def build_prerequisites(row):
    eligibility = clean(row.get("Eligibility"))
    if eligibility:
        return eligibility
    fallback_parts = [
        clean(row.get("Target Audience")),
        clean(row.get("Target Audiences (resolved)")),
        clean(row.get("Category")),
        clean(row.get("Industry")),
        clean(row.get("Sector")),
    ]
    fallback = "; ".join(unique(fallback_parts))
    if fallback:
        return f"Eligibility not specified separately in source file. Likely applicable to: {fallback}."
    return "Eligibility not specified separately in source file. Refer scheme details and official scheme link before application."


def build_tags(row, geographies):
    seed = [
        "financial schemes",
        "financial support",
        "loans and schemes",
        "money",
        row.get("Category"),
        row.get("Category (API)"),
        row.get("Vertical"),
        row.get("Primary Domain"),
        row.get("Secondary Domain"),
        row.get("Target Audience"),
        row.get("Gender"),
        row.get("Sector"),
        row.get("Industry"),
        row.get("Name of Ministry"),
        row.get("Scheme Provider"),
    ]
    tags = []
    for item in seed:
        tags.extend(split_terms(item))
    tags.extend(split_terms(row.get("Keywords")))
    tags.extend(geographies)
    return unique(tags)[:40]


def map_row(row):
    geographies, geography_warnings = normalize_geographies(row)
    title = clean(row.get("Scheme Title"))
    provider = clean(row.get("Scheme Provider")) or clean(row.get("Implementing Agency")) or clean(row.get("Name of Ministry")) or ORG_NAME
    description = build_description(row)
    tags = build_tags(row, geographies)
    scheme_info = clean_url(row.get("Link to scheme information")) or clean_url(row.get("Scheme URL (wep.gov.in resource-detail)"))
    enrollment = clean_url(row.get("Link for enrollment"))
    government = clean_url(row.get("Link to government website"))
    notes = []
    if geography_warnings:
        notes.extend(geography_warnings)
    if not provider or provider == ORG_NAME:
        notes.append("Provider fallback used")
    if not scheme_info and not enrollment and not government:
        notes.append("No source/action link available")
    if not clean(row.get("Eligibility")):
        notes.append("Eligibility fallback used")

    record = {
        "SourceSchemeId": clean(row.get("Scheme ID")),
        "SourceSlug": clean(row.get("Slug")),
        "TraderId": TRADER_ID,
        "OrganisationName": ORG_NAME,
        "SolutionName": title,
        "OfferingName": title,
        "OfferingCategory": DEFAULT_CATEGORY,
        "OfferingGroup": DEFAULT_GROUP,
        "OfferingType": DEFAULT_TYPE,
        "6M": DEFAULT_6M,
        "PrimaryValuechain": DEFAULT_VALUECHAIN,
        "PrimaryApplication": DEFAULT_APPLICATION,
        "Valuechains": DEFAULT_VALUECHAIN,
        "Applications": DEFAULT_APPLICATION,
        "Tags": "; ".join(tags),
        "Languages": DEFAULT_LANGUAGE,
        "Geographies": "; ".join(geographies),
        "AboutSolution": description,
        "AboutOffering": description,
        "Who Can avail it": clean(row.get("Target Audience")) or clean(row.get("Target Audiences (resolved)")) or "Entrepreneurs, MSMEs, SHGs, Producer groups, Organisations",
        "Trainer Name": provider,
        "Trainer Email Address": "",
        "Trainer Phone Number": "",
        "Trainer Details": build_trainer_details(row),
        "Duration": "1 month",
        "Prerequisites - Participants and Training": build_prerequisites(row),
        "Cost (Service)": "Can be quoted after finalising scope",
        "Support post Service": "Yes, if in location",
        "Support post Service Cost": "No",
        "Is it offered - Online or Offline": infer_delivery_mode(row),
        "Certification Offered": "No",
        "Remarks on Cost": clean(row.get("Benefits")),
        "Location Availability": "; ".join(geographies),
        "Service offering Brochure": scheme_info,
        "Contact Details": "\n".join(
            value for value in [
                f"Provider: {provider}" if provider else "",
                f"Enrollment: {enrollment}" if enrollment else "",
                f"Scheme information: {scheme_info}" if scheme_info else "",
                f"Government website: {government}" if government else "",
            ]
            if value
        ),
        "Offering Link on GRE": "",
        "SourceSection": clean(row.get("Section")),
        "SourceProvider": provider,
        "SourceGovernmentLink": government,
        "SourceSchemeInfoLink": scheme_info,
        "SourceEnrollmentLink": enrollment,
        "ImportNotes": "; ".join(unique(notes)),
    }
    raw = {str(key): clean(value) for key, value in row.items()}
    return record, raw


def validation_summary(records):
    missing_required = []
    for record in records:
        missing = [field for field in ["SourceSchemeId", "OfferingName", "AboutOffering", "Geographies", "Prerequisites - Participants and Training"] if not record.get(field)]
        if missing:
            missing_required.append({"scheme_id": record.get("SourceSchemeId"), "title": record.get("OfferingName"), "missing": missing})

    titles = Counter(record["OfferingName"].strip().lower() for record in records if record.get("OfferingName"))
    scheme_ids = Counter(record["SourceSchemeId"] for record in records if record.get("SourceSchemeId"))
    duplicate_titles = {title: count for title, count in titles.items() if count > 1}
    duplicate_scheme_ids = {scheme_id: count for scheme_id, count in scheme_ids.items() if count > 1}
    notes = Counter(note.strip() for record in records for note in record.get("ImportNotes", "").split(";") if note.strip())
    geographies = Counter(geo.strip() for record in records for geo in record.get("Geographies", "").split(";") if geo.strip())
    delivery_modes = Counter(record.get("Is it offered - Online or Offline", "") for record in records)
    providers = Counter(record.get("SourceProvider", "") for record in records if record.get("SourceProvider"))
    sections = Counter(record.get("SourceSection", "") for record in records if record.get("SourceSection"))

    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total_records": len(records),
        "ready_records": len(records) - len(missing_required),
        "missing_required": missing_required[:100],
        "missing_required_count": len(missing_required),
        "duplicate_title_count": len(duplicate_titles),
        "duplicate_titles_sample": dict(list(duplicate_titles.items())[:50]),
        "duplicate_scheme_id_count": len(duplicate_scheme_ids),
        "duplicate_scheme_ids": duplicate_scheme_ids,
        "import_notes": dict(notes.most_common(50)),
        "top_geographies": dict(geographies.most_common(40)),
        "delivery_modes": dict(delivery_modes),
        "top_providers": dict(providers.most_common(25)),
        "sections": dict(sections),
    }


def write_csv(path, records):
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=IMPORT_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)


def write_markdown(path, summary, csv_name, json_name):
    lines = [
        "# Financial Schemes Import Dry Run",
        "",
        f"Generated: `{summary['generated_at']}`",
        "",
        "## Outputs",
        "",
        f"- Import preview CSV: `{csv_name}`",
        f"- Full preview JSON: `{json_name}`",
        "",
        "## Record Counts",
        "",
        f"- Total source rows mapped: `{summary['total_records']}`",
        f"- Rows ready by required-field check: `{summary['ready_records']}`",
        f"- Rows missing required fields: `{summary['missing_required_count']}`",
        f"- Duplicate title groups: `{summary['duplicate_title_count']}`",
        f"- Duplicate Scheme ID groups: `{summary['duplicate_scheme_id_count']}`",
        "",
        "## Fixed Mapping",
        "",
        "- Organisation / Trader: `GRE Curated Financing Solutions` / `738`",
        "- Offering Category: `Service offerings`",
        "- Offering Group: `Service`",
        "- Offering Type: `Financial Support`",
        "- 6M: `Money`",
        "- Value Chain: `Financial`",
        "- Application: `Loans and Schemes`",
        "- Language: `English`",
        "- Duration: `1 month`",
        "- Cost: `Can be quoted after finalising scope`",
        "- Support Post Service: `Yes, if in location`",
        "- Support Post Service Cost: `No`",
        "- Certification: `No`",
        "",
        "## Top Geographies",
        "",
    ]
    for key, count in summary["top_geographies"].items():
        lines.append(f"- `{key}`: {count}")
    lines.extend(["", "## Delivery Modes", ""])
    for key, count in summary["delivery_modes"].items():
        lines.append(f"- `{key}`: {count}")
    lines.extend(["", "## Source Sections", ""])
    for key, count in summary["sections"].items():
        lines.append(f"- `{key}`: {count}")
    lines.extend(["", "## Top Providers", ""])
    for key, count in summary["top_providers"].items():
        lines.append(f"- `{key}`: {count}")
    lines.extend(["", "## Import Notes / Warnings", ""])
    if summary["import_notes"]:
        for key, count in summary["import_notes"].items():
            lines.append(f"- `{key}`: {count}")
    else:
        lines.append("- No warnings generated.")
    if summary["missing_required"]:
        lines.extend(["", "## Missing Required Field Sample", ""])
        for item in summary["missing_required"][:25]:
            lines.append(f"- `{item['scheme_id']}` `{item['title']}` missing: {', '.join(item['missing'])}")
    lines.extend([
        "",
        "## Recommended Next Step",
        "",
        "Review the CSV in Excel and spot-check the schemes that have `ImportNotes`. After approval, use this mapped file as the source for a Supabase import action that inserts into `solutions` and `offerings` with the same fields.",
    ])
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def sql_string(value):
    if value is None:
        return "null"
    return "'" + str(value).replace("'", "''") + "'"


def html_from_text(text):
    blocks = [block.strip() for block in clean(text).split("\n\n") if block.strip()]
    if not blocks:
        return ""
    return "\n".join(f"<p>{block}</p>" for block in blocks)


def search_document(record):
    parts = [
        record.get("SolutionName"),
        record.get("OfferingName"),
        record.get("OfferingCategory"),
        record.get("OfferingGroup"),
        record.get("OfferingType"),
        record.get("6M"),
        record.get("PrimaryValuechain"),
        record.get("PrimaryApplication"),
        record.get("Tags"),
        record.get("Languages"),
        record.get("Geographies"),
        record.get("AboutOffering"),
        record.get("Trainer Name"),
        record.get("Trainer Details"),
        record.get("Contact Details"),
    ]
    return re.sub(r"\s+", " ", " | ".join(clean(part) for part in parts if clean(part))).strip()


def to_json_array(value):
    return split_terms(value) if isinstance(value, str) else unique(value or [])


def import_payload(records, raw_rows):
    payload = []
    for record, raw in zip(records, raw_rows):
        scheme_id = record["SourceSchemeId"]
        solution_id = f"MIS-SOLUTION-FINSCHEME-{scheme_id}"
        offering_id = f"MIS-OFFERING-FINSCHEME-{scheme_id}"
        raw_payload = {
            "source": "financial_schemes_excel_import",
            "source_scheme_id": scheme_id,
            "source_slug": record["SourceSlug"],
            "source_section": record["SourceSection"],
            "source_provider": record["SourceProvider"],
            "source_links": {
                "wep_resource": raw.get("Scheme URL (wep.gov.in resource-detail)", ""),
                "government": record["SourceGovernmentLink"],
                "scheme_information": record["SourceSchemeInfoLink"],
                "enrollment": record["SourceEnrollmentLink"],
            },
            "mapped_by": "prepare_financial_schemes_import.py",
            "mapped_at": datetime.now(timezone.utc).isoformat(),
            "raw": raw,
        }
        item = {
            "source_scheme_id": scheme_id,
            "solution_id": solution_id,
            "offering_id": offering_id,
            "trader_id": record["TraderId"],
            "solution_name": record["SolutionName"],
            "offering_name": record["OfferingName"],
            "solution_status": "Approved in MIS",
            "publish_status": "MIS Published",
            "created_at_source": datetime.now(timezone.utc).isoformat(),
            "about_solution_html": html_from_text(record["AboutSolution"]),
            "about_solution_text": record["AboutSolution"],
            "about_offering_html": html_from_text(record["AboutOffering"]),
            "about_offering_text": record["AboutOffering"],
            "offering_category": record["OfferingCategory"],
            "offering_group": record["OfferingGroup"],
            "offering_type": record["OfferingType"],
            "domain_6m": record["6M"],
            "primary_valuechain": record["PrimaryValuechain"],
            "primary_application": record["PrimaryApplication"],
            "valuechains": to_json_array(record["Valuechains"]),
            "applications": to_json_array(record["Applications"]),
            "tags": to_json_array(record["Tags"]),
            "languages": to_json_array(record["Languages"]),
            "geographies": split_terms(record["Geographies"]),
            "geographies_raw": record["Geographies"],
            "audience": record["Who Can avail it"],
            "trainer_name": record["Trainer Name"],
            "trainer_email": record["Trainer Email Address"],
            "trainer_phone": record["Trainer Phone Number"],
            "trainer_details_html": html_from_text(record["Trainer Details"]),
            "trainer_details_text": record["Trainer Details"],
            "duration": record["Duration"],
            "prerequisites": record["Prerequisites - Participants and Training"],
            "service_cost": record["Cost (Service)"],
            "support_post_service": record["Support post Service"],
            "support_post_service_cost": record["Support post Service Cost"],
            "delivery_mode": record["Is it offered - Online or Offline"],
            "certification_offered": record["Certification Offered"],
            "cost_remarks": record["Remarks on Cost"],
            "location_availability": record["Location Availability"],
            "service_brochure_url": record["Service offering Brochure"],
            "contact_details": record["Contact Details"],
            "gre_link": record["Offering Link on GRE"],
            "search_document": search_document(record),
            "source_row_signature": hashlib.sha256(json.dumps(raw, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest(),
            "raw_payload": raw_payload,
        }
        payload.append(item)
    return payload


def write_import_sql(path, payload):
    payload_json = json.dumps(payload, ensure_ascii=False)
    sql = f"""begin;

with payload as (
  select *
  from jsonb_to_recordset({sql_string(payload_json)}::jsonb) as p(
    source_scheme_id text,
    solution_id text,
    offering_id text,
    trader_id text,
    solution_name text,
    offering_name text,
    solution_status text,
    publish_status text,
    created_at_source text,
    about_solution_html text,
    about_solution_text text,
    about_offering_html text,
    about_offering_text text,
    offering_category text,
    offering_group text,
    offering_type text,
    domain_6m text,
    primary_valuechain text,
    primary_application text,
    valuechains jsonb,
    applications jsonb,
    tags jsonb,
    languages jsonb,
    geographies jsonb,
    geographies_raw text,
    audience text,
    trainer_name text,
    trainer_email text,
    trainer_phone text,
    trainer_details_html text,
    trainer_details_text text,
    duration text,
    prerequisites text,
    service_cost text,
    support_post_service text,
    support_post_service_cost text,
    delivery_mode text,
    certification_offered text,
    cost_remarks text,
    location_availability text,
    service_brochure_url text,
    contact_details text,
    gre_link text,
    search_document text,
    source_row_signature text,
    raw_payload jsonb
  )
),
inserted_solutions as (
  insert into public.solutions (
    solution_id,
    trader_id,
    solution_name,
    solution_status,
    publish_status,
    created_at_source,
    about_solution_html,
    about_solution_text,
    solution_image_url,
    raw_payload
  )
  select
    p.solution_id,
    p.trader_id,
    p.solution_name,
    p.solution_status,
    p.publish_status,
    p.created_at_source,
    p.about_solution_html,
    p.about_solution_text,
    null,
    p.raw_payload
  from payload p
  where not exists (
    select 1
    from public.solutions s
    where s.solution_id = p.solution_id
       or ((s.raw_payload ->> 'source') = 'financial_schemes_excel_import'
           and (s.raw_payload ->> 'source_scheme_id') = p.source_scheme_id)
  )
  returning solution_id
),
inserted_offerings as (
  insert into public.offerings (
    offering_id,
    solution_id,
    trader_id,
    publish_status,
    created_at_source,
    offering_name,
    offering_category,
    offering_group,
    offering_type,
    domain_6m,
    primary_valuechain_id,
    primary_valuechain,
    primary_application_id,
    primary_application,
    valuechains,
    applications,
    tags,
    languages,
    geographies,
    geographies_raw,
    about_offering_html,
    about_offering_text,
    audience,
    trainer_name,
    trainer_email,
    trainer_phone,
    trainer_details_html,
    trainer_details_text,
    duration,
    prerequisites,
    service_cost,
    support_post_service,
    support_post_service_cost,
    delivery_mode,
    certification_offered,
    cost_remarks,
    location_availability,
    service_brochure_url,
    grade_capacity,
    product_cost,
    lead_time,
    support_details,
    product_brochure_url,
    knowledge_content_url,
    contact_details,
    gre_link,
    search_document,
    last_import_id,
    raw_payload,
    source_row_signature,
    rule_thematic_hints,
    rule_service_hints,
    rule_keywords,
    rule_6m_signals,
    ai_thematic_area,
    ai_application_area,
    ai_offering_kind,
    ai_service_kind,
    ai_keywords,
    ai_6m_signals,
    ai_summary,
    ai_engine,
    ai_enriched_at,
    ai_enrichment_status,
    ai_payload,
    ai_prompt_version,
    ai_schema_version
  )
  select
    p.offering_id,
    p.solution_id,
    p.trader_id,
    p.publish_status,
    p.created_at_source,
    p.offering_name,
    p.offering_category,
    p.offering_group,
    p.offering_type,
    p.domain_6m,
    null,
    p.primary_valuechain,
    null,
    p.primary_application,
    array(select jsonb_array_elements_text(p.valuechains)),
    array(select jsonb_array_elements_text(p.applications)),
    array(select jsonb_array_elements_text(p.tags)),
    array(select jsonb_array_elements_text(p.languages)),
    array(select jsonb_array_elements_text(p.geographies)),
    p.geographies_raw,
    p.about_offering_html,
    p.about_offering_text,
    p.audience,
    p.trainer_name,
    nullif(p.trainer_email, ''),
    nullif(p.trainer_phone, ''),
    p.trainer_details_html,
    p.trainer_details_text,
    p.duration,
    p.prerequisites,
    p.service_cost,
    p.support_post_service,
    p.support_post_service_cost,
    p.delivery_mode,
    p.certification_offered,
    nullif(p.cost_remarks, ''),
    p.location_availability,
    nullif(p.service_brochure_url, ''),
    null,
    null,
    null,
    null,
    null,
    null,
    p.contact_details,
    nullif(p.gre_link, ''),
    p.search_document,
    null,
    p.raw_payload,
    p.source_row_signature,
    array['Financial', 'Loans and Schemes'],
    array['Financial Support'],
    array(select jsonb_array_elements_text(p.tags)),
    array['Money'],
    'Financial',
    'Loans and Schemes',
    'service',
    'financial support',
    array(select jsonb_array_elements_text(p.tags)),
    array['Money'],
    left(p.about_offering_text, 500),
    'rules',
    now(),
    'ready',
    jsonb_build_object(
      'source', 'financial_schemes_excel_import',
      'source_scheme_id', p.source_scheme_id,
      'mapping', jsonb_build_object(
        'offering_category', p.offering_category,
        'offering_type', p.offering_type,
        'domain_6m', p.domain_6m,
        'primary_valuechain', p.primary_valuechain,
        'primary_application', p.primary_application
      )
    ),
    'financial-schemes-import-v1',
    'financial-schemes-import-v1'
  from payload p
  where not exists (
    select 1
    from public.offerings o
    where o.offering_id = p.offering_id
       or ((o.raw_payload ->> 'source') = 'financial_schemes_excel_import'
           and (o.raw_payload ->> 'source_scheme_id') = p.source_scheme_id)
  )
  returning offering_id
)
select
  (select count(*) from payload) as source_rows,
  (select count(*) from inserted_solutions) as inserted_solutions,
  (select count(*) from inserted_offerings) as inserted_offerings;

commit;
"""
    path.write_text(sql, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Prepare a dry-run import package for financial schemes.")
    parser.add_argument("--source", required=True, help="Path to Financial Schemes.xlsx")
    parser.add_argument("--output-dir", required=True, help="Directory for dry-run outputs")
    parser.add_argument("--sql-output", help="Optional SQL file for an idempotent Supabase import")
    parser.add_argument("--sql-batch-dir", help="Optional directory for batched idempotent Supabase import SQL files")
    parser.add_argument("--sql-batch-size", type=int, default=75, help="Rows per batched SQL file")
    args = parser.parse_args()

    source = Path(args.source)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    frame = pd.read_excel(source, sheet_name="All Schemes")
    records = []
    raw_rows = []
    preview = []
    for _, row in frame.iterrows():
        mapped, raw = map_row(row.to_dict())
        records.append(mapped)
        raw_rows.append(raw)
        preview.append({"mapped": mapped, "raw": raw})

    csv_path = output_dir / "financial_schemes_import_preview.csv"
    json_path = output_dir / "financial_schemes_import_preview.json"
    summary_path = output_dir / "financial_schemes_validation_report.md"
    summary_json_path = output_dir / "financial_schemes_validation_summary.json"

    write_csv(csv_path, records)
    json_path.write_text(json.dumps(preview, ensure_ascii=False, indent=2), encoding="utf-8")
    summary = validation_summary(records)
    summary_json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    write_markdown(summary_path, summary, csv_path.name, json_path.name)
    sql_path = None
    import_rows = import_payload(records, raw_rows)
    if args.sql_output:
      sql_path = Path(args.sql_output)
      sql_path.parent.mkdir(parents=True, exist_ok=True)
      write_import_sql(sql_path, import_rows)
    batch_paths = []
    if args.sql_batch_dir:
      batch_dir = Path(args.sql_batch_dir)
      batch_dir.mkdir(parents=True, exist_ok=True)
      batch_size = max(1, args.sql_batch_size)
      for index in range(0, len(import_rows), batch_size):
          batch_number = (index // batch_size) + 1
          batch_path = batch_dir / f"financial_schemes_import_batch_{batch_number:03d}.sql"
          write_import_sql(batch_path, import_rows[index:index + batch_size])
          batch_paths.append(str(batch_path))

    print(json.dumps({
        "csv": str(csv_path),
        "json": str(json_path),
        "summary": str(summary_path),
        "summary_json": str(summary_json_path),
        "sql": str(sql_path) if sql_path else None,
        "sql_batches": batch_paths,
        "total_records": summary["total_records"],
        "ready_records": summary["ready_records"],
        "missing_required_count": summary["missing_required_count"],
        "duplicate_title_count": summary["duplicate_title_count"],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
