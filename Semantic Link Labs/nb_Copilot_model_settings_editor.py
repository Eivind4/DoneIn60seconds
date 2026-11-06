#!/usr/bin/env python
# coding: utf-8

# ## nb_share_model_settings_editor
# 
# Shareable, self-contained Copilot model settings editor (Python notebook)

# # Copilot settings editor for semantic models
# 
# Turn **Copilot and MCP tools for readers** and **Approved for Copilot** On/Off for your semantic models, and see per model whether it is **prepped for AI** (AI instructions, AI schema, verified answers or a data agent).
# 
# **How to use:** set the parameters, *Run all*, pick **On / Off** per model in the editor, press **Preview**, then **Apply**.
# 
# - ⚠ = the model is available for Copilot/MCP but has none of the AI preparation.
# - **AI schema** counts only when it exposes less than 95% of the visible tables, columns and measures (MCP returns the whole model when no schema is set) - same rule as the AI readiness solution. See *AI schema coverage* under **Show status**.
# - Changing settings needs **write permission** on the model; reading AI preparation needs **Build permission** and the tenant setting *Users can use the Power BI Model Context Protocol server endpoint*.
# - Uses the internal (undocumented) endpoint behind the model settings page, which can change without notice - test on one model first.
# 

# In[ ]:


# The command is not a standard IPython magic command. It is designed for use within Fabric notebooks only.
# %pip install -q semantic-link


# In[ ]:


# Workspace name/ID, a list of them, or None = every workspace you can access
workspace      = None
include_models = None   # model names/IDs to show, e.g. ["Sales", "Finance"]. None = all
skip_models    = None   # model names to hide, e.g. ["Sales (test)"]


# In[ ]:


import base64
import json
import re
import time

import notebookutils
import pandas as pd
import requests
import sempy.fabric as fabric

DEFAULT_SKIP_MODELS = ["ModelBPA", "Fabric Capacity Metrics", "AI readiness model"]
POWERBI_MCP_URL = "https://api.fabric.microsoft.com/v1/mcp/powerbi"

# Bits in model.modelProperties (verified by toggling in the portal):
#   8 set  = "Approved for Copilot" On
#   32 set = "Copilot and MCP tools for readers" Off (inverted)
APPROVED_BIT, RESTRICT_READERS_BIT = 8, 32

READERS, APPROVED = "Copilot & MCP for readers", "Approved for Copilot"
PREP_ITEMS = ["AI instructions", "AI schema", "Verified answers", "Data agent"]

_CLUSTER: dict = {}
_MODEL_META: dict = {}


def _as_list(value) -> list:
    if value is None:
        return []
    return [value] if isinstance(value, str) else list(value)


def _headers() -> dict:
    # sempy's default credential gets 401 on the internal endpoints, the notebook token works
    return {"Authorization": "Bearer " + notebookutils.credentials.getToken("pbi")}


def _check(resp, hint_403: str = "needs write permission on the model"):
    if not resp.ok:
        hint = f" ({hint_403})" if resp.status_code in (401, 403) else ""
        raise RuntimeError(f"HTTP {resp.status_code}{hint}")


def _cluster(ws_id) -> str:
    """Backend cluster of the workspace, e.g. https://wabi-west-europe-redirect.analysis.windows.net."""
    if ws_id not in _CLUSTER:
        r = requests.get(f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets", headers=_headers(), timeout=60)
        r.raise_for_status()
        _CLUSTER[ws_id] = r.json()["@odata.context"].split("/v1.0")[0]
    return _CLUSTER[ws_id]


def list_models(workspace=None, include_models=None, skip_models=None) -> pd.DataFrame:
    """Semantic models in scope, excluding the default models of lakehouses/warehouses."""
    wanted = {str(w).strip().lower() for w in _as_list(workspace)}
    include = {str(m).strip().lower() for m in _as_list(include_models)}
    skip = set(DEFAULT_SKIP_MODELS) | set(_as_list(skip_models))

    dfW = fabric.list_workspaces("type ne 'AdminInsights'")
    if wanted:
        dfW = dfW[dfW["Name"].str.lower().isin(wanted) | dfW["Id"].astype(str).str.lower().isin(wanted)]
    if dfW.empty:
        raise ValueError("No matching workspaces found.")

    rows = []
    for _, w in dfW.iterrows():
        try:
            dfD = fabric.list_datasets(workspace=w["Id"], mode="rest")
            items = fabric.list_items(workspace=w["Id"])
        except Exception as e:
            print(f"⚠ Workspace '{w['Name']}' skipped: {str(e)[:120]}")
            continue
        default_models = set(items.loc[items["Type"].isin(["Lakehouse", "Warehouse"]), "Display Name"])
        for _, d in dfD.iterrows():
            name, ds_id = d["Dataset Name"], str(d["Dataset Id"])
            if name in skip or name in default_models:
                continue
            if include and name.strip().lower() not in include and ds_id.lower() not in include:
                continue
            rows.append({"Workspace": w["Name"], "Semantic model": name,
                         "Workspace Id": str(w["Id"]), "Dataset Id": ds_id})
    return pd.DataFrame(rows, columns=["Workspace", "Semantic model", "Workspace Id", "Dataset Id"])


def read_settings(ws_id, ds_id) -> dict:
    r = requests.get(f"{_cluster(ws_id)}/metadata/models/{ds_id}", headers=_headers(), timeout=60)
    _check(r)
    model = r.json().get("model") or {}
    _MODEL_META[ds_id] = model
    bits = int(model.get("modelProperties") or 0)
    return {READERS: not bits & RESTRICT_READERS_BIT, APPROVED: bool(bits & APPROVED_BIT)}


def write_setting(ws_id, ds_id, setting: str, on: bool) -> bool:
    """Changes one setting and returns True when the new value is read back."""
    headers, cluster = _headers(), _cluster(ws_id)
    if setting == READERS:
        # The settings POST needs the internal model id, not the item GUID
        model_id = (_MODEL_META.get(ds_id) or {}).get("id") or ds_id
        r = requests.post(f"{cluster}/metadata/models/{model_id}/settings", headers=headers,
                          json={"restrictCopilotForReaders": not on}, timeout=60)
    else:
        r = requests.put(f"{cluster}/metadata/workspace/{ws_id}/model/{ds_id}/preppedCopilot", headers=headers,
                         json={"preppedForCopilot": on, "isReadOnly": False}, timeout=60)
    _check(r)
    for _ in range(6):  # the service can apply the change with a delay
        if read_settings(ws_id, ds_id)[setting] == on:
            return True
        time.sleep(3)
    return False


AI_SCHEMA_EXPOSE_ALL_PCT = 95   # same rule as the AI readiness solution: coverage >= 95% = no real AI schema


def _visible_model_objects(ws_id, ds_id) -> dict:
    """Visible tables, plus visible columns/measures per visible table (auto date tables excluded)."""
    def visible(df):
        col = next((c for c in ("Is Hidden", "Hidden", "Measure Hidden", "Column Hidden") if c in df.columns), None)
        return df if col is None else df[df[col].astype(str).str.lower() != "true"]

    tables = visible(fabric.list_tables(dataset=ds_id, workspace=ws_id))["Name"].astype(str)
    tables = set(tables[~tables.str.startswith(("DateTableTemplate_", "LocalDateTable_"))])
    cols = visible(fabric.list_columns(dataset=ds_id, workspace=ws_id))
    measures = visible(fabric.list_measures(dataset=ds_id, workspace=ws_id))
    by_table = lambda df, name_col: {t: set(g[name_col].astype(str)) for t, g in df[df["Table Name"].isin(tables)].groupby("Table Name")}
    return {"tables": tables, "columns": by_table(cols, "Column Name"), "measures": by_table(measures, "Measure Name")}


def _object_keys(tables, columns: dict, measures: dict) -> set:
    return (set(tables)
            | {f"'{t}'[{c}]" for t, cs in columns.items() for c in cs}
            | {f"'{t}'[{m}]" for t, ms in measures.items() for m in ms})


def ai_schema_coverage(schema_tables, model: dict) -> tuple:
    """(objects in AI schema, visible objects in model) - only visible model objects are counted.
    A schema table without listed columns/measures exposes all of them (as in the AI readiness solution)."""
    def names(items):
        return {str(i.get("Name") or i.get("name") if isinstance(i, dict) else i) for i in items or []} - {"None", ""}

    tables, columns, measures = set(), {}, {}
    for t in schema_tables or []:
        name = str(t.get("Name") or t.get("name") or "") if isinstance(t, dict) else str(t)
        if name not in model["tables"]:
            continue
        tables.add(name)
        listed_cols = names(t.get("Columns") or t.get("columns")) if isinstance(t, dict) else set()
        listed_measures = names(t.get("Measures") or t.get("measures")) if isinstance(t, dict) else set()
        columns[name] = listed_cols or model["columns"].get(name, set())
        measures[name] = listed_measures or model["measures"].get(name, set())
    visible = _object_keys(model["tables"], model["columns"], model["measures"])
    return len(_object_keys(tables, columns, measures) & visible), len(visible)


def read_ai_prep(ws_id, ds_id) -> dict:
    """AI instructions / AI schema / verified answers via the Power BI MCP server (GetSemanticModelSchema)."""
    payload = {"jsonrpc": "2.0", "id": 1, "method": "tools/call",
               "params": {"name": "GetSemanticModelSchema", "arguments": {"artifactId": ds_id}}}
    r = requests.post(POWERBI_MCP_URL, json=payload, timeout=120, headers={
        **_headers(), "Content-Type": "application/json", "Accept": "application/json, text/event-stream"})
    _check(r, "needs Build permission and the Power BI MCP tenant setting")
    msg = next((json.loads(line[5:]) for line in r.text.splitlines() if line.startswith("data:")), None) or r.json()
    result = msg.get("result") or {}
    if "error" in msg or result.get("isError"):
        raise RuntimeError(str(msg.get("error") or result.get("content"))[:200])
    parsed = json.loads(result["content"][0]["text"])
    schema = parsed.get("schema", parsed)

    # With no AI schema set, MCP returns the whole visible model, so a schema only counts when it is narrower
    try:
        in_schema, visible = ai_schema_coverage(schema.get("Tables"), _visible_model_objects(ws_id, ds_id))
        pct = round(100 * in_schema / visible) if visible else None
        has_schema = None if pct is None else (in_schema > 0 and pct < AI_SCHEMA_EXPOSE_ALL_PCT)
        coverage = f"{in_schema} of {visible} objects ({pct}%)" if pct is not None else "?"
    except Exception:
        has_schema, coverage = None, "? (model objects not readable)"
    return {
        "AI instructions": bool(str(schema.get("CustomInstructions") or "").strip()),
        "AI schema": has_schema,
        "AI schema coverage": coverage,
        "Verified answers": bool(schema.get("VerifiedAnswers")),
    }


def _fabric_api(method: str, path: str):
    """Fabric REST call that follows long-running operations."""
    client = fabric.FabricRestClient()
    resp = client.post(path) if method == "POST" else client.get(path)
    if resp.status_code == 202:
        location = resp.headers.get("Location")
        for _ in range(60):
            time.sleep(int(resp.headers.get("Retry-After", 2) or 2))
            status = (client.get(location).json() or {}).get("status")
            if status == "Succeeded":
                return client.get(location.rstrip("/") + "/result").json()
            if status in ("Failed", "Cancelled"):
                raise RuntimeError(f"Operation {status}")
        raise TimeoutError("Long-running operation did not finish in time.")
    _check(resp, "no access")
    return resp.json() if resp.content else {}


def models_with_data_agent(ws_id) -> set:
    """Lower-case IDs of the semantic models used by data agents in the workspace."""
    ids, path = set(), f"v1/workspaces/{ws_id}/items?type=DataAgent"
    agents = []
    while path:
        body = _fabric_api("GET", path)
        agents += body.get("value") or []
        path = body.get("continuationUri")
    for agent in agents:
        parts = (_fabric_api("POST", f"v1/workspaces/{ws_id}/items/{agent['id']}/getDefinition")
                 .get("definition") or {}).get("parts") or []
        for p in parts:
            if not re.match(r"^Files/Config/(draft|published)/[^/]+/datasource\.json$", p.get("path") or ""):
                continue
            ds = json.loads(base64.b64decode(p.get("payload") or "").decode("utf-8", errors="replace") or "{}")
            if "semantic_model" in str(ds.get("type", "")).lower():
                ids.add(str(ds.get("artifactId", "")).lower())
    return ids


def is_prepped(row: dict):
    """True if any AI preparation exists, False if none, None if unknown."""
    values = [row.get(k) for k in PREP_ITEMS]
    return True if any(v is True for v in values) else (False if all(v is False for v in values) else None)


def needs_warning(row: dict, readers=None, approved=None) -> bool:
    """Available for Copilot/MCP (current or planned values) without any AI preparation."""
    readers = row[READERS] if readers is None else readers
    approved = row[APPROVED] if approved is None else approved
    return bool(readers or approved) and is_prepped(row) is False


def status_table(state: list) -> pd.DataFrame:
    mark = {True: "✓", False: "✗", None: "?"}
    return pd.DataFrame([{
        "Workspace": r["Workspace"], "Semantic model": r["Semantic model"],
        READERS: {True: "On", False: "Off", None: "?"}[r[READERS]],
        APPROVED: {True: "On", False: "Off", None: "?"}[r[APPROVED]],
        **{k: mark[r[k]] for k in PREP_ITEMS},
        "AI schema coverage": r.get("AI schema coverage") or "?",
        "Warning": "⚠ Available for Copilot/MCP, not prepped for AI" if needs_warning(r) else "",
    } for r in state])


# ### Read current settings and AI preparation

# In[ ]:


models = list_models(workspace, include_models, skip_models)
print(f"Reading {len(models)} model(s)...")

problems = {"settings": [], "ai prep": [], "data agents": []}
agent_models = set()
for ws_id, ws_name in models[["Workspace Id", "Workspace"]].drop_duplicates().values:
    try:
        agent_models |= models_with_data_agent(ws_id)
    except Exception as e:
        problems["data agents"].append(f"{ws_name}: {str(e)[:80]}")

state = []
for _, m in models.iterrows():
    has_agent = m["Dataset Id"].lower() in agent_models or (None if problems["data agents"] else False)
    row = {**m.to_dict(), READERS: None, APPROVED: None, "AI instructions": None, "AI schema": None,
           "Verified answers": None, "Data agent": has_agent}
    try:
        row.update(read_settings(m["Workspace Id"], m["Dataset Id"]))
    except Exception as e:
        problems["settings"].append(f"{m['Workspace']} / {m['Semantic model']}: {str(e)[:80]}")
    try:
        row.update(read_ai_prep(m["Workspace Id"], m["Dataset Id"]))
    except Exception as e:
        problems["ai prep"].append(f"{m['Workspace']} / {m['Semantic model']}: {str(e)[:80]}")
    state.append(row)

titles = {"settings": "Copilot settings not readable/changeable (write permission needed)",
          "ai prep": "AI preparation not readable (Build permission / MCP tenant setting needed)",
          "data agents": "Data agents not readable in workspace"}
for key, items in problems.items():
    if items:
        print(f"⚠ {titles[key]}:\n  - " + "\n  - ".join(items))
print(f"Done. {sum(needs_warning(r) for r in state)} model(s) available for Copilot/MCP without AI preparation.")


# ### Editor
# Pick **On / Off** per model (*Keep* = no change). **Preview** shows what will change, **Apply** changes it and shows the result, **Show status** lists which settings are on and the AI preparation per model.

# In[ ]:


import ipywidgets as widgets
from IPython.display import display, clear_output


def _prep_text(row: dict) -> str:
    if is_prepped(row) is None:
        return "? (not readable)"
    found = [k for k in PREP_ITEMS if row[k]]
    if row["AI schema"] is None and row.get("AI schema coverage"):
        found.append("AI schema ?")
    return ", ".join(found) if found else "None"


def model_settings_editor(state: list):
    def dropdown(current, readable):
        now = {True: "On", False: "Off", None: "?"}[current]
        return widgets.Dropdown(options=[(f"Keep ({now})", "Keep"), ("On", "On"), ("Off", "Off")], value="Keep",
                                disabled=not readable, layout=widgets.Layout(width="15%"))

    def col(text, width, bold=False):
        return widgets.HTML(f"<b>{text}</b>" if bold else text, layout=widgets.Layout(width=width))

    header = widgets.HBox([col("Workspace", "16%", True), col("Semantic model", "20%", True),
                           col("Prepped for AI", "26%", True), col("", "3%"),
                           col(READERS, "15%", True), col(APPROVED, "15%", True)])
    rows, controls = [], []
    for r in state:
        readable = r[READERS] is not None
        readers, approved = dropdown(r[READERS], readable), dropdown(r[APPROVED], readable)
        warn = col("", "3%")

        def refresh_warning(_=None, r=r, readers=readers, approved=approved, warn=warn):
            planned = lambda dd, cur: cur if dd.value == "Keep" else dd.value == "On"
            show = needs_warning(r, planned(readers, r[READERS]), planned(approved, r[APPROVED]))
            warn.value = "<span title='Available for Copilot/MCP, but not prepped for AI'>⚠</span>" if show else ""

        readers.observe(refresh_warning, "value")
        approved.observe(refresh_warning, "value")
        refresh_warning()
        controls.append((r, readers, approved))
        rows.append(widgets.HBox([
            col(r["Workspace"], "16%"),
            col(r["Semantic model"] + ("" if readable else " <i>(no permission)</i>"), "20%"),
            col(_prep_text(r), "26%"), warn, readers, approved,
        ]))

    preview = widgets.Button(description="Preview", icon="eye", button_style="info")
    apply = widgets.Button(description="Apply", icon="check", button_style="danger")
    output = widgets.Output()

    def planned_changes() -> list:
        changes = []
        for r, readers, approved in controls:
            for setting, dd in ((READERS, readers), (APPROVED, approved)):
                if dd.value != "Keep" and (dd.value == "On") != r[setting]:
                    changes.append((r, setting, dd.value == "On"))
        return changes

    def run(do_apply: bool):
        with output:
            clear_output()
            changes = planned_changes()
            if not changes:
                print("Nothing to change.")
                return
            results = []
            for r, setting, on in changes:
                result = {"Workspace": r["Workspace"], "Semantic model": r["Semantic model"], "Setting": setting,
                          "From": "On" if r[setting] else "Off", "To": "On" if on else "Off"}
                if do_apply:
                    try:
                        ok = write_setting(r["Workspace Id"], r["Dataset Id"], setting, on)
                        result["Status"] = "Changed" if ok else "Sent, but not confirmed - check the model settings"
                        r[setting] = on if ok else r[setting]
                    except Exception as e:
                        result["Status"] = f"Failed: {str(e)[:120]}"
                results.append(result)
            display(pd.DataFrame(results))
            if do_apply:
                for _, readers, approved in controls:
                    readers.value = approved.value = "Keep"
                print("Current status:")
                display(status_table(state))

    def show_status(_):
        with output:
            clear_output()
            display(status_table(state))

    show = widgets.Button(description="Show status", icon="table")
    preview.on_click(lambda _: run(False))
    apply.on_click(lambda _: run(True))
    show.on_click(show_status)
    display(widgets.VBox([header, *rows, widgets.HBox([preview, apply, show]), output]))


model_settings_editor(state)

