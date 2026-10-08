#!/usr/bin/env python3
"""Assets Export People enriched with cascaded Team -> Theme -> Stream + Role (v2) + STEP 6 OnlyContributors rows.

Enrichment rules
----------------
People objectTypeId=35

A) Team cascade
- From People attribute "Team member of" (reference -> Team objects, objectTypeId=34)
  enrich the following columns from the referenced Team objects:
    Team Description            <- Team."Description"
    Team Email                  <- Team."Team email" (alias: Team Email)
    Team Channels               <- Team."Team Channels"
    Team Development Manager    <- Team."Development Manager"
    Team Contributors           <- Team."Contributors"
    Team Solution Architect     <- Team."Architect" (alias: Solution Architect)
    Team Delivery Manager       <- Team."Delivery Manager"
    Team Security Champion      <- Team."Security Champion"
    Team Managed Services       <- Team."Managed Services"
    Team Page                   <- Team."Team Page"
    Team Scrum Master           <- Team."Scrum Master"
    Team Product Manager        <- Team."Product Manager"
    Team Service Manager        <- Team."Service Manager"

- From the referenced Team object, read Team."Theme" and:
    Team Theme                  <- Team."Theme" (label(s))

- Then resolve each Theme (objectTypeId=32) referenced by Team Theme and:
    Team Theme Description      <- Theme."Description"
    Team Stream                 <- Theme."Stream" (label(s))

- Then resolve each Stream (objectTypeId=33) referenced by Team Stream and:
    Team Stream Description     <- Stream."Description"

B) Role cascade
- From People attribute "Role" (reference -> Role objects, objectTypeId=36)
    Role Description            <- Role."Description"
    Role Grants                 <- Role."Role Grants"

Multi-values
-----------
All multi-values are de-duplicated and joined with '||'.

CUSTOM RULE (STEP 6 - 2026-05)
------------------------------
After generating the CSV, we scan all Teams and find those with:
- zero direct members (no People referencing the team in "Team member of")
- ONLY Contributors (no other team guest roles set)

For each of these teams we append one synthetic row:
- Name = "OnlyContributors#<n>"
- Team member of = that team
- Team Contributors = all contributors (|| joined)
- Email/User/Key are set (if columns exist) to a unique placeholder value
The JS will ignore these rows as "people" but will still use Team Contributors
to populate the team in the org chart.
"""

import os
import csv
import time
import base64
import argparse
import re
import requests
from typing import Any, Dict, List, Optional


SITE = os.environ.get("ATLASSIAN_SITE", "https://instance.atlassian.net").rstrip("/")
WORKSPACE_ID = os.environ["ASSETS_WORKSPACE_ID"]
EMAIL = os.environ["ATLASSIAN_EMAIL"]
API_TOKEN = os.environ["ATLASSIAN_API_TOKEN"]

BASE = f"{SITE}/gateway/api/jsm/assets/workspace/{WORKSPACE_ID}/v1"
_auth = base64.b64encode(f"{EMAIL}:{API_TOKEN}".encode("utf-8")).decode("utf-8")
HEADERS = {
    "Authorization": f"Basic {_auth}",
    "Accept": "application/json",
    "Content-Type": "application/json",
}

TIMEOUT = (10, 60)


def req(method: str, url: str, payload: Optional[dict], max_retries: int, backoff: float) -> Any:
    last = None
    for attempt in range(1, max_retries + 1):
        try:
            if method == 'GET':
                r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            else:
                r = requests.post(url, headers=HEADERS, json=payload, timeout=TIMEOUT)

            if r.ok:
                try:
                    return r.json()
                except Exception:
                    return r.text

            try:
                body = r.json()
            except Exception:
                body = (r.text or '')[:2000]
            raise RuntimeError(f"HTTP {r.status_code} {url}: {body}")
        except Exception as e:
            last = e
            if attempt < max_retries:
                time.sleep(backoff * attempt)
            else:
                break
    raise RuntimeError(f"Failed after {max_retries} retries: {method} {url}: {last}")


def get_objecttype_attributes(object_type_id: str, max_retries: int, backoff: float) -> List[dict]:
    data = req('GET', f"{BASE}/objecttype/{object_type_id}/attributes", None, max_retries, backoff)
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        vals = data.get('values')
        return vals if isinstance(vals, list) else []
    raise RuntimeError(f"Unexpected attributes response type: {type(data)}")


def post_totalcount(ql_query: str, max_retries: int, backoff: float) -> Optional[int]:
    try:
        data = req('POST', f"{BASE}/object/aql/totalcount", {"qlQuery": ql_query}, max_retries, backoff)
    except Exception:
        return None
    if isinstance(data, int):
        return data
    if isinstance(data, dict):
        for k in ('count', 'total', 'totalCount', 'value'):
            v = data.get(k)
            if isinstance(v, int):
                return v
    return None


def post_aql(ql_query: str, max_results: int, include_attributes: bool, max_retries: int, backoff: float) -> List[dict]:
    payload = {
        "qlQuery": ql_query,
        "startAt": 0,
        "maxResults": max_results,
        "includeAttributes": include_attributes,
    }
    data = req('POST', f"{BASE}/object/aql", payload, max_retries, backoff)
    if isinstance(data, dict):
        return data.get('values') or data.get('objectEntries') or []
    return []


def build_exclusion_clause(ids: List[str], chunk_size: int = 50) -> str:
    if not ids:
        return ''
    parts = []
    for i in range(0, len(ids), chunk_size):
        chunk = ids[i:i+chunk_size]
        parts.append(f" AND objectId NOT IN ({','.join(chunk)})")
    return ''.join(parts)


def fetch_all_objects_excluding(object_type_id: str,
                                batch_size: int,
                                include_attributes: bool,
                                max_retries: int,
                                backoff: float,
                                sleep_sec: float) -> List[dict]:
    base_query = f"objectTypeId = {object_type_id}"
    total = post_totalcount(base_query, max_retries, backoff)
    print(f"[fetch] objectTypeId={object_type_id} totalcount={total}", flush=True)

    seen_ids: List[str] = []
    seen_set = set()
    out: List[dict] = []

    while True:
        q = base_query + build_exclusion_clause(seen_ids, chunk_size=50)
        objs = post_aql(q, batch_size, include_attributes, max_retries, backoff)

        new_objs = []
        for o in objs:
            oid = str(o.get('id'))
            if oid and oid not in seen_set:
                new_objs.append(o)

        if not new_objs:
            break

        for o in new_objs:
            oid = str(o.get('id'))
            seen_set.add(oid)
            seen_ids.append(oid)
            out.append(o)

        if total is not None:
            print(f"  progress {len(seen_ids)}/{total}", flush=True)
            if len(seen_ids) >= total:
                break
        else:
            print(f"  progress {len(seen_ids)}", flush=True)

        if sleep_sec:
            time.sleep(sleep_sec)

    return out


def resolve_attr_id(attr_defs: List[dict], aliases: List[str]) -> Optional[str]:
    alias_set = {a.strip().lower() for a in aliases if a and a.strip()}
    for a in attr_defs:
        name = a.get('name')
        if isinstance(name, str) and name.strip().lower() in alias_set:
            return str(a.get('id'))
    return None


def values_from_attr_id(obj: dict, attr_id: str, *, split_commas: bool = True) -> List[str]:
    vals: List[str] = []
    for a in obj.get('attributes', []) or []:
        if str(a.get('objectTypeAttributeId')) != str(attr_id):
            continue
        for v in a.get('objectAttributeValues', []) or []:
            ref = v.get('referencedObject')
            if ref:
                lbl = ref.get('label') or ref.get('objectKey')
                if lbl:
                    vals.append(str(lbl))
            else:
                if v.get('value') is not None:
                    vals.append(str(v.get('value')))

    out: List[str] = []
    for x in vals:
        if not x:
            continue
        s = str(x).strip().replace('\r\n', '\n').replace('\r', '\n')
        s = s.replace('\n,', '\n').replace(',\n', '\n')
        if '||' in s:
            parts = s.split('||')
        elif '\n' in s:
            parts = s.split('\n')
        elif split_commas and ',' in s:
            parts = s.split(',')
        else:
            parts = [s]
        for p in parts:
            p = p.strip().strip('"').lstrip(',').strip()
            if p:
                out.append(p)

    seen = set()
    dedup: List[str] = []
    for v in out:
        if v not in seen:
            seen.add(v)
            dedup.append(v)
    return dedup


def load_template_header(path: str) -> Optional[List[str]]:
    if not path:
        return None
    try:
        with open(path, newline='', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            header = next(reader)
        return [h.strip() for h in header if h and h.strip()]
    except Exception:
        return None


def build_columns(base_attr_names: List[str], template_path: str, extra_cols: List[str]) -> List[str]:
    tpl = load_template_header(template_path)
    cols: List[str] = []
    seen = set()

    if tpl:
        for h in tpl:
            if h in base_attr_names and h not in seen:
                cols.append(h); seen.add(h)
            elif h in extra_cols and h not in seen:
                cols.append(h); seen.add(h)

    for h in base_attr_names:
        if h not in seen:
            cols.append(h); seen.add(h)

    for e in extra_cols:
        if e not in seen:
            cols.append(e); seen.add(e)

    return cols


def join_dedup(values: List[str]) -> str:
    seen = set()
    out = []
    for v in values:
        v = (v or '').strip()
        if not v:
            continue
        if v not in seen:
            seen.add(v)
            out.append(v)
    return '||'.join(out)


def flatten_object(obj: dict, attr_def_by_id: Dict[str, dict], attr_names: List[str]) -> Dict[str, str]:
    row = {c: '' for c in attr_names}
    for a in obj.get('attributes', []) or []:
        aid = str(a.get('objectTypeAttributeId'))
        ad = attr_def_by_id.get(aid)
        if not ad:
            continue
        name = ad.get('name', aid)
        if name not in row:
            continue

        vals: List[str] = []
        for v in a.get('objectAttributeValues', []) or []:
            ref = v.get('referencedObject')
            if ref:
                vals.append(str(ref.get('label') or ref.get('objectKey') or ''))
            else:
                if v.get('value') is not None:
                    vals.append(str(v.get('value')))
        row[name] = '||'.join([x for x in vals if x])
    return row


def main():
    parser = argparse.ArgumentParser(description='Export People enriched with Team->Theme->Stream and Role cascades (v2) + STEP 6')
    parser.add_argument('--out', default='people_enriched.csv')
    parser.add_argument('--template', default='')

    parser.add_argument('--people-type', default='35')
    parser.add_argument('--team-type', default='34')
    parser.add_argument('--role-type', default='36')
    parser.add_argument('--theme-type', default='32')
    parser.add_argument('--stream-type', default='33')

    parser.add_argument('--batch-size', type=int, default=25)
    parser.add_argument('--max-retries', type=int, default=5)
    parser.add_argument('--backoff', type=float, default=1.5)
    parser.add_argument('--sleep', type=float, default=0.1)

    args = parser.parse_args()

    people_type = str(args.people_type)
    team_type = str(args.team_type)
    role_type = str(args.role_type)
    theme_type = str(args.theme_type)
    stream_type = str(args.stream_type)

    print(f"Site: {SITE}")
    print(f"Workspace: {WORKSPACE_ID}")

    # --- Resolve attribute IDs ---
    people_attr_defs = get_objecttype_attributes(people_type, args.max_retries, args.backoff)
    people_team_member_id = resolve_attr_id(people_attr_defs, ['Team member of', 'Team Member Of', 'Team'])
    people_role_id = resolve_attr_id(people_attr_defs, ['Role', 'Roles', 'User Role'])

    if not people_team_member_id:
        print("[WARN] Could not resolve People 'Team member of' attribute id. Team enrichment will be empty.")
    if not people_role_id:
        print("[WARN] Could not resolve People 'Role' attribute id. Role enrichment will be empty.")

    team_attr_defs = get_objecttype_attributes(team_type, args.max_retries, args.backoff)
    # Team fields
    team_desc_id = resolve_attr_id(team_attr_defs, ['Description', 'Team Description'])
    team_email_id = resolve_attr_id(team_attr_defs, ['Team email', 'Team Email', 'Email'])
    team_channels_id = resolve_attr_id(team_attr_defs, ['Team Channels', 'Channels'])
    team_devmgr_id = resolve_attr_id(team_attr_defs, ['Development Manager', 'Dev Manager', 'Development manager'])
    team_contr_id = resolve_attr_id(team_attr_defs, ['Contributors', 'Team Contributors'])
    team_arch_id = resolve_attr_id(team_attr_defs, ['Architect', 'Solution Architect', 'Team Solution Architect'])
    team_delivery_id = resolve_attr_id(team_attr_defs, ['Delivery Manager', 'Team Delivery Manager'])
    team_secchamp_id = resolve_attr_id(team_attr_defs, ['Security Champion', 'Team Security Champion'])
    team_managed_services_id = resolve_attr_id(team_attr_defs, ['Managed Services', 'Managed services', 'Managed'])
    team_page_id = resolve_attr_id(team_attr_defs, ['Team Page', 'Page', 'Confluence Page'])
    team_scrum_id = resolve_attr_id(team_attr_defs, ['Scrum Master', 'Team Scrum Master'])
    team_pm_id = resolve_attr_id(team_attr_defs, ['Product Manager', 'Team Product Manager', 'Team Product Manager'])
    team_sm_id = resolve_attr_id(team_attr_defs, ['Service Manager', 'Team Service Manager'])

    # Team -> Theme
    team_theme_id = resolve_attr_id(team_attr_defs, ['Theme', 'Team Theme', 'Product Theme'])

    theme_attr_defs = get_objecttype_attributes(theme_type, args.max_retries, args.backoff)
    theme_desc_id = resolve_attr_id(theme_attr_defs, ['Description', 'Theme Description', 'Team Theme Description'])
    theme_stream_id = resolve_attr_id(theme_attr_defs, ['Stream', 'Team Stream'])

    stream_attr_defs = get_objecttype_attributes(stream_type, args.max_retries, args.backoff)
    stream_desc_id  = resolve_attr_id(stream_attr_defs, ['Description', 'Stream Description', 'Team Stream Description'])
    stream_boost_id = resolve_attr_id(stream_attr_defs, ['Visual Boost'])

    team_boost_id  = resolve_attr_id(team_attr_defs,  ['Visual Boost'])
    theme_boost_id = resolve_attr_id(theme_attr_defs, ['Visual Boost'])

    role_attr_defs = get_objecttype_attributes(role_type, args.max_retries, args.backoff)
    role_desc_id = resolve_attr_id(role_attr_defs, ['Description', 'Role Description'])
    role_grants_id = resolve_attr_id(role_attr_defs, ['Role Grants', 'Grants'])

    # --- Load caches ---
    print("\n[STEP 1] Loading Teams (typeId=34)...", flush=True)
    teams = fetch_all_objects_excluding(team_type, args.batch_size, True, args.max_retries, args.backoff, args.sleep)
    print(f"Loaded {len(teams)} teams.", flush=True)

    teams_by_label: Dict[str, dict] = {}
    for t in teams:
        lbl = str(t.get('label') or '').strip()
        if lbl:
            teams_by_label[lbl.lower()] = t

    print("\n[STEP 2] Loading Themes (typeId=32)...", flush=True)
    themes = fetch_all_objects_excluding(theme_type, args.batch_size, True, args.max_retries, args.backoff, args.sleep)
    print(f"Loaded {len(themes)} themes.", flush=True)

    themes_by_label: Dict[str, dict] = {}
    for th in themes:
        lbl = str(th.get('label') or '').strip()
        if lbl:
            themes_by_label[lbl.lower()] = th

    print("\n[STEP 3] Loading Streams (typeId=33)...", flush=True)
    streams = fetch_all_objects_excluding(stream_type, args.batch_size, True, args.max_retries, args.backoff, args.sleep)
    print(f"Loaded {len(streams)} streams.", flush=True)

    streams_by_label: Dict[str, dict] = {}
    for st in streams:
        lbl = str(st.get('label') or '').strip()
        if lbl:
            streams_by_label[lbl.lower()] = st

    print("\n[STEP 4] Loading Roles (typeId=36)...", flush=True)
    roles = fetch_all_objects_excluding(role_type, args.batch_size, True, args.max_retries, args.backoff, args.sleep)
    print(f"Loaded {len(roles)} roles.", flush=True)

    roles_by_label: Dict[str, dict] = {}
    for r in roles:
        lbl = str(r.get('label') or '').strip()
        if lbl:
            roles_by_label[lbl.lower()] = r

    print("\n[STEP 5] Loading People (typeId=35)...", flush=True)
    people = fetch_all_objects_excluding(people_type, args.batch_size, True, args.max_retries, args.backoff, args.sleep)
    print(f"Loaded {len(people)} people.", flush=True)

    # Base people columns
    people_attr_names = [a.get('name', f"attr_{a.get('id')}") for a in people_attr_defs]
    people_attr_by_id = {str(a.get('id')): a for a in people_attr_defs if a.get('id') is not None}

    # Real column name for the People attribute "Team member of"
    people_team_member_name = None
    if people_team_member_id:
        people_team_member_name = people_attr_by_id.get(str(people_team_member_id), {}).get('name')

    extra_cols = [
        'Team Description',
        'Team Email',
        'Team Channels',
        'Team Development Manager',
        'Team Contributors',
        'Team Solution Architect',
        'Team Delivery Manager',
        'Team Security Champion',
        'Team Managed Services',
        'Team Page',
        'Team Scrum Master',
        'Team Product Manager',
        'Team Service Manager',
        'Team Theme',
        'Team Theme Description',
        'Team Stream',
        'Team Stream Description',
        'Team Visual Boost',
        'Team Theme Visual Boost',
        'Team Stream Visual Boost',
        'Role Description',
        'Role Grants',
    ]

    columns = build_columns(people_attr_names, args.template, extra_cols)

    def enrich_from_teams(team_labels: List[str]) -> Dict[str, str]:
        agg = {
            'Team Description': [],
            'Team Email': [],
            'Team Channels': [],
            'Team Development Manager': [],
            'Team Contributors': [],
            'Team Solution Architect': [],
            'Team Delivery Manager': [],
            'Team Security Champion': [],
            'Team Managed Services': [],
            'Team Page': [],
            'Team Scrum Master': [],
            'Team Product Manager': [],
            'Team Service Manager': [],
            'Team Theme': [],
            'Team Theme Description': [],
            'Team Stream': [],
            'Team Stream Description': [],
            'Team Visual Boost': [],
            'Team Theme Visual Boost': [],
            'Team Stream Visual Boost': [],
        }

        for tl in team_labels:
            t = teams_by_label.get(tl.strip().lower())
            if not t:
                continue

            if team_desc_id:
                agg['Team Description'].extend(values_from_attr_id(t, team_desc_id, split_commas=False))
            if team_email_id:
                agg['Team Email'].extend(values_from_attr_id(t, team_email_id))
            if team_channels_id:
                agg['Team Channels'].extend(values_from_attr_id(t, team_channels_id))
            if team_devmgr_id:
                agg['Team Development Manager'].extend(values_from_attr_id(t, team_devmgr_id))
            if team_contr_id:
                agg['Team Contributors'].extend(values_from_attr_id(t, team_contr_id))
            if team_arch_id:
                agg['Team Solution Architect'].extend(values_from_attr_id(t, team_arch_id))
            if team_delivery_id:
                agg['Team Delivery Manager'].extend(values_from_attr_id(t, team_delivery_id))
            if team_secchamp_id:
                agg['Team Security Champion'].extend(values_from_attr_id(t, team_secchamp_id))
            if team_managed_services_id:
                agg['Team Managed Services'].extend(values_from_attr_id(t, team_managed_services_id))
            if team_page_id:
                agg['Team Page'].extend(values_from_attr_id(t, team_page_id))
            if team_scrum_id:
                agg['Team Scrum Master'].extend(values_from_attr_id(t, team_scrum_id))
            if team_pm_id:
                agg['Team Product Manager'].extend(values_from_attr_id(t, team_pm_id))
            if team_sm_id:
                agg['Team Service Manager'].extend(values_from_attr_id(t, team_sm_id))
            if team_boost_id:
                agg['Team Visual Boost'].extend(values_from_attr_id(t, team_boost_id))

            # Team -> Theme -> Stream
            theme_labels: List[str] = []
            if team_theme_id:
                theme_labels = values_from_attr_id(t, team_theme_id)
                agg['Team Theme'].extend(theme_labels)

            for th_label in theme_labels:
                th = themes_by_label.get(th_label.strip().lower())
                if not th:
                    continue

                if theme_desc_id:
                    agg['Team Theme Description'].extend(values_from_attr_id(th, theme_desc_id, split_commas=False))
                if theme_boost_id:
                    agg['Team Theme Visual Boost'].extend(values_from_attr_id(th, theme_boost_id))

                stream_labels: List[str] = []
                if theme_stream_id:
                    stream_labels = values_from_attr_id(th, theme_stream_id)
                    agg['Team Stream'].extend(stream_labels)

                for st_label in stream_labels:
                    st = streams_by_label.get(st_label.strip().lower())
                    if not st:
                        continue
                    if stream_desc_id:
                        agg['Team Stream Description'].extend(values_from_attr_id(st, stream_desc_id, split_commas=False))
                    if stream_boost_id:
                        agg['Team Stream Visual Boost'].extend(values_from_attr_id(st, stream_boost_id))

        return {k: join_dedup(v) for k, v in agg.items()}

    def enrich_from_roles(role_labels: List[str]) -> Dict[str, str]:
        desc_vals: List[str] = []
        grants_vals: List[str] = []

        for rl in role_labels:
            r = roles_by_label.get(rl.strip().lower())
            if not r:
                continue
            if role_desc_id:
                desc_vals.extend(values_from_attr_id(r, role_desc_id, split_commas=False))
            if role_grants_id:
                grants_vals.extend(values_from_attr_id(r, role_grants_id))

        return {
            'Role Description': join_dedup(desc_vals),
            'Role Grants': join_dedup(grants_vals),
        }

    def norm(s: str) -> str:
        return ' '.join((s or '').strip().lower().split())

    def slug(s: str) -> str:
        s = (s or '').strip().lower()
        s = re.sub(r'[^a-z0-9]+', '-', s)
        return s.strip('-') or 'x'

    # ------------------------------
    # Pre-compute direct membership per team (from People.Team member of)
    # ------------------------------
    team_direct_count: Dict[str, int] = {}
    if people_team_member_id:
        for p in people:
            direct_teams = values_from_attr_id(p, people_team_member_id, split_commas=False)
            for tl in direct_teams:
                k = tl.strip().lower()
                if not k:
                    continue
                team_direct_count[k] = team_direct_count.get(k, 0) + 1

    # Canonical team label map
    team_label_canon: Dict[str, str] = {}
    for tl_lower, t in teams_by_label.items():
        team_label_canon[tl_lower] = (str(t.get('label') or '').strip() or tl_lower)

    # Helper: team has ONLY contributors (no other guest roles)
    other_role_attr_ids = [
        team_devmgr_id, team_arch_id, team_delivery_id, team_secchamp_id,
        team_scrum_id, team_pm_id, team_sm_id
    ]
    other_role_attr_ids = [x for x in other_role_attr_ids if x]
    # Helper: team has ANY guest population (contributors OR any other guest role)
    guest_population_attr_ids = [team_contr_id] + other_role_attr_ids
    guest_population_attr_ids = [x for x in guest_population_attr_ids if x]

    def team_has_any_guest_population(team_obj: dict) -> bool:
        for aid in guest_population_attr_ids:
            if values_from_attr_id(team_obj, aid):
                return True
        return False


    # Helper: team has a valid Theme -> Stream cascade
    def team_has_valid_theme_stream(team_obj: dict) -> bool:
        if not team_theme_id or not theme_stream_id:
            return False

        theme_labels = values_from_attr_id(team_obj, team_theme_id)
        if not theme_labels:
            return False

        for th_label in theme_labels:
            th = themes_by_label.get(th_label.strip().lower())
            if not th:
                continue

            stream_labels = values_from_attr_id(th, theme_stream_id)
            if not stream_labels:
                continue

            for st_label in stream_labels:
                if streams_by_label.get(st_label.strip().lower()):
                    return True

        return False

    # ------------------------------
    # Write CSV
    # ------------------------------
    with open(args.out, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()

        # (A) Normal people rows
        for p in people:
            base = flatten_object(p, people_attr_by_id, people_attr_names)
            row = {c: '' for c in columns}
            for k, v in base.items():
                if k in row:
                    row[k] = v

            # Teams (DIRECT only)
            team_labels: List[str] = []
            if people_team_member_id:
                team_labels = values_from_attr_id(p, people_team_member_id, split_commas=False)
            row.update(enrich_from_teams(team_labels))

            # Roles
            role_labels: List[str] = []
            if people_role_id:
                role_labels = values_from_attr_id(p, people_role_id)
            row.update(enrich_from_roles(role_labels))

            writer.writerow(row)

        # (B) STEP 6 - OnlyContributors synthetic rows
        print("\n[STEP 6] Generating OnlyContributors rows for teams with 0 direct members and only Contributors...", flush=True)

        only_idx = 0
        for tl_lower, team_obj in teams_by_label.items():
            # zero direct members
            if team_direct_count.get(tl_lower, 0) > 0:
                continue

            # must have contributors OR other guest roles
            if not team_has_any_guest_population(team_obj):
                continue
        
            # must have valid Theme -> Stream cascade
            if not team_has_valid_theme_stream(team_obj):
                continue

            only_idx += 1
            team_label = team_label_canon.get(tl_lower, tl_lower)
            team_key = slug(team_label)
            oc_name = f"OnlyContributors#{only_idx}"

            row = {c: '' for c in columns}

            # Minimal identity fields (avoid "fully empty key" drop in JS)
            if 'Name' in row:
                row['Name'] = oc_name
            if 'Status' in row and not row['Status']:
                row['Status'] = 'Active'
            if 'Email' in row:
                row['Email'] = f"onlycontributors+{only_idx}.{team_key}@invalid.local"
            if 'User' in row and not row['User']:
                row['User'] = f"onlycontributors-{only_idx}-{team_key}"
            if 'Key' in row and not row['Key']:
                row['Key'] = f"ONLYCONTRIB-{only_idx}-{team_key}"

            # Team membership (use real attribute name if present)
            if people_team_member_name and people_team_member_name in row:
                row[people_team_member_name] = team_label
            elif 'Team member of' in row:
                row['Team member of'] = team_label

            # Enrich from this team to get Stream/Theme metadata + other team fields
            row.update(enrich_from_teams([team_label]))

            # Override Team Contributors with ALL contributors
            contribs = values_from_attr_id(team_obj, team_contr_id) if team_contr_id else []
            if contribs and 'Team Contributors' in row:
                row['Team Contributors'] = join_dedup(contribs)

            writer.writerow(row)

        print(f"[STEP 6] Added {only_idx} OnlyContributors rows.", flush=True)

    print(f"✅ Export completed: {args.out}", flush=True)


if __name__ == '__main__':
    main()