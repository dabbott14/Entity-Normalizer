"""Entity Normalizer — Streamlit dashboard.

Run with:
    streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from normalizer import NormalizerSession, suggest_clusters
from normalizer.suggest import BACKEND

SAMPLE = Path(__file__).parent / "data" / "sample_case.xlsx"

st.set_page_config(page_title="Entity Normalizer", page_icon="🧹", layout="wide")


# ----------------------------------------------------------------------
# State
# ----------------------------------------------------------------------
def _state():
    ss = st.session_state
    ss.setdefault("session", None)
    ss.setdefault("source_name", None)
    ss.setdefault("selected", [])      # match keys queued for renaming
    ss.setdefault("new_name", "")
    ss.setdefault("flash", None)       # (kind, message) shown once after a rerun
    ss.setdefault("table_version", 0)  # bump to clear table selections
    return ss


ss = _state()


def flash(kind: str, message: str) -> None:
    ss.flash = (kind, message)


def load(source, name: str, rename_all_rows: bool) -> None:
    session = NormalizerSession.from_excel(source, rename_all_rows=rename_all_rows)
    ss.session, ss.source_name = session, name
    ss.selected, ss.new_name = [], ""
    ss.table_version += 1
    msg = f"Loaded **{name}**."
    if session.restored_steps:
        msg += f" Restored {session.restored_steps} rename step(s) from its Alias Table."
    flash("success", msg)


# ----------------------------------------------------------------------
# Callbacks (run before the page redraws, so they may set widget values)
# ----------------------------------------------------------------------
def add_to_selection(keys: list[str]) -> None:
    for k in keys:
        if k not in ss.selected:
            ss.selected.append(k)
    # Prefill the new-name box with the most common spelling, like
    # double-clicking a name in the original form.
    if not ss.new_name.strip():
        ss.new_name = ss.session.suggested_name(ss.selected)
    ss.table_version += 1


def remove_from_selection(keys: list[str]) -> None:
    ss.selected = [k for k in ss.selected if k not in keys]
    if not ss.selected:
        ss.new_name = ""
    ss.table_version += 1


def use_as_new_name(name: str) -> None:
    ss.new_name = name


def confirm_rename(names_by_key: dict[str, str]) -> None:
    try:
        step = ss.session.rename([names_by_key.get(k, k) for k in ss.selected], ss.new_name)
    except ValueError as exc:
        flash("warning", str(exc))
        return
    flash("success", f"Renamed to **{step.target}** — {step.cells_changed} cell(s) updated.")
    ss.selected, ss.new_name = [], ""
    ss.table_version += 1


def undo() -> None:
    step = ss.session.undo()
    if step:
        flash("info", f"Undid rename to **{step.target}** ({step.cells_changed} cell(s) restored).")
    else:
        flash("warning", "Nothing to undo.")
    ss.table_version += 1


def apply_cluster(names: list[str], key: str) -> None:
    target = ss.get(key, "").strip()
    try:
        step = ss.session.rename(names, target)
    except ValueError as exc:
        flash("warning", str(exc))
        return
    flash("success", f"Merged {len(names)} names into **{step.target}** — {step.cells_changed} cell(s) updated.")
    ss.table_version += 1


def send_cluster(names: list[str], suggested: str) -> None:
    from normalizer import match_key

    ss.selected = []
    ss.new_name = suggested
    add_to_selection([match_key(n) for n in names])


# ----------------------------------------------------------------------
# Sidebar: data source and settings
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("Case data")
    rename_all_rows = st.toggle(
        "Include non-wire rows on Combined",
        value=False,
        help="Off: only Combined rows where Tab Name = Wire are listed and renamed. "
        "On: every Combined row is included (e.g. check payees).",
    )
    upload = st.file_uploader("Upload a workbook (.xlsx)", type=["xlsx"])
    c1, c2 = st.columns(2)
    if c1.button("Load upload", disabled=upload is None, width="stretch"):
        load(upload, upload.name, rename_all_rows)
    if c2.button("Load sample", disabled=not SAMPLE.exists(), width="stretch"):
        load(SAMPLE, SAMPLE.name, rename_all_rows)

    if ss.session is not None:
        st.divider()
        st.subheader("Reuse an alias table")
        alias_file = st.file_uploader(
            "Alias table (.csv with Original, Normalized)", type=["csv"], key="alias_upload"
        )
        if st.button("Apply alias table", disabled=alias_file is None, width="stretch"):
            try:
                added = ss.session.apply_alias_table(pd.read_csv(alias_file))
                flash("success", f"Applied {added} rename step(s) from the alias table.")
                ss.table_version += 1
            except ValueError as exc:
                flash("warning", str(exc))
            st.rerun()

    with st.expander("How to use this tool"):
        st.markdown(
            "1. Load a workbook with **Wire** and **Combined** sheets.\n"
            "2. Search the **Entity List** and tick one or more names.\n"
            "3. Click **Add →** to queue them under **Entities to Rename**.\n"
            "4. Check the **New name** box (it prefills with the most common "
            "spelling) and click **Confirm rename**.\n"
            "5. **Undo** reverses the last rename and can be repeated.\n"
            "6. The **Suggested matches** tab groups likely variants for one-click merging.\n"
            "7. **Download** the normalized workbook; it includes the mapping log "
            "and an alias table, and reopening it restores your work."
        )

st.title("Entity Normalizer")
st.caption("Standardize counterparty names across wire activity — every change logged and reversible.")

if ss.flash:
    kind, message = ss.flash
    getattr(st, kind)(message)
    ss.flash = None

session: NormalizerSession | None = ss.session
if session is None:
    st.info("Load a workbook from the sidebar to begin. **Load sample** opens a synthetic case.")
    st.stop()

if session.rename_all_rows != rename_all_rows:
    session.rename_all_rows = rename_all_rows
    ss.table_version += 1

for w in session.warnings:
    st.warning(w)

names = session.distinct_names()
names_by_key = dict(zip(names["key"], names["Name"]))
log = session.mapping_log()

m1, m2, m3, m4 = st.columns(4)
m1.metric("Distinct names", len(names))
m2.metric("Rename steps", len(log))
m3.metric("Cells normalized", int(log["Cells changed"].sum()) if len(log) else 0)
m4.metric("Workbook", ss.source_name or "—")

tab_norm, tab_suggest, tab_log, tab_preview = st.tabs(
    ["Normalize", "Suggested matches", "Mapping log", "Preview data"]
)

# ----------------------------------------------------------------------
# Normalize: the three-panel layout from the original form
# ----------------------------------------------------------------------
with tab_norm:
    left, middle, queue, right = st.columns([6, 1.2, 4, 3.5], gap="medium")

    with left:
        st.markdown("##### Entity List")
        search = st.text_input("Search", placeholder="Search for a name (optional)…", label_visibility="collapsed")
        view = names
        if search.strip():
            view = names[names["Name"].str.contains(search.strip(), case=False, regex=False)]
        view = view[~view["key"].isin(ss.selected)].reset_index(drop=True)
        event = st.dataframe(
            view[["Name", "Count", "Spellings", "Appears in"]],
            hide_index=True,
            width="stretch",
            height=420,
            on_select="rerun",
            selection_mode="multi-row",
            key=f"entity_table_{ss.table_version}",
            column_config={
                "Count": st.column_config.NumberColumn(width="small"),
                "Spellings": st.column_config.NumberColumn(
                    width="small", help="Distinct spellings merged into this entry (case/spacing)"
                ),
            },
        )
        picked_left = view.loc[event.selection.rows, "key"].tolist() if event.selection.rows else []
        st.caption(f"{len(view)} of {len(names)} names shown")

    with middle:
        st.markdown("<div style='height:9rem'></div>", unsafe_allow_html=True)
        st.button(
            "Add →", on_click=add_to_selection, args=(picked_left,),
            disabled=not picked_left, width="stretch",
        )

    with queue:
        st.markdown("##### Entities to Rename")
        queued = names[names["key"].isin(ss.selected)].reset_index(drop=True)
        q_event = st.dataframe(
            queued[["Name", "Count"]],
            hide_index=True,
            width="stretch",
            height=330,
            on_select="rerun",
            selection_mode="multi-row",
            key=f"queue_table_{ss.table_version}",
        )
        picked_queue = queued.loc[q_event.selection.rows, "key"].tolist() if q_event.selection.rows else []
        b1, b2 = st.columns(2)
        b1.button(
            "← Remove", on_click=remove_from_selection, args=(picked_queue,),
            disabled=not picked_queue, width="stretch",
        )
        b2.button(
            "Use as new name",
            on_click=use_as_new_name,
            args=(names_by_key.get(picked_queue[0], "") if len(picked_queue) == 1 else "",),
            disabled=len(picked_queue) != 1,
            width="stretch",
            help="Copy the selected name into the New name box to edit it",
        )
        if len(queued):
            st.caption(f"{len(queued)} name(s), {int(queued['Count'].sum())} cell(s) queued")

    with right:
        st.markdown("##### Actions")
        st.text_input(
            "New name",
            key="new_name",
            disabled=not ss.selected,
            placeholder="Add names to rename first",
        )
        st.button(
            "Confirm rename",
            type="primary",
            on_click=confirm_rename,
            args=(names_by_key,),
            disabled=not ss.selected or not ss.new_name.strip(),
            width="stretch",
        )
        st.button("Undo last rename", on_click=undo, disabled=len(log) == 0, width="stretch")
        st.divider()
        stem = Path(ss.source_name or "case").stem
        st.download_button(
            "Download normalized workbook",
            data=session.to_excel_bytes(),
            file_name=f"{stem}_normalized.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

# ----------------------------------------------------------------------
# Suggested matches
# ----------------------------------------------------------------------
with tab_suggest:
    threshold = st.slider(
        "Match sensitivity (minimum similarity)", 75, 100, 88,
        help="Lower finds more possible matches but more false positives.",
    )
    try:
        clusters = suggest_clusters(names["Name"], names["Count"], threshold=threshold)
    except ValueError as exc:
        clusters = []
        st.warning(str(exc))

    st.caption(
        f"{len(clusters)} group(s) of likely variants, largest first. "
        f"Nothing is changed until you apply a group. (matching engine: {BACKEND})"
    )
    for i, c in enumerate(clusters):
        with st.container(border=True):
            top, actions = st.columns([3, 2])
            with top:
                st.markdown(
                    f"**{c['suggested']}** · {len(c['names'])} spellings · "
                    f"{c['total']} cells · similarity ≥ {c['score']:.0f}"
                )
                st.dataframe(
                    pd.DataFrame({"Name": c["names"], "Count": c["counts"]}),
                    hide_index=True, width="stretch",
                )
            with actions:
                key = f"cluster_name_{ss.table_version}_{i}"
                st.text_input("Merge into", value=c["suggested"], key=key)
                a1, a2 = st.columns(2)
                a1.button(
                    "Apply", key=f"apply_{ss.table_version}_{i}", type="primary",
                    on_click=apply_cluster, args=(c["names"], key), width="stretch",
                )
                a2.button(
                    "Review in panel", key=f"send_{ss.table_version}_{i}",
                    on_click=send_cluster, args=(c["names"], c["suggested"]),
                    width="stretch",
                    help="Queue these names on the Normalize tab to adjust before renaming",
                )

# ----------------------------------------------------------------------
# Mapping log
# ----------------------------------------------------------------------
with tab_log:
    if log.empty:
        st.info("No renames yet.")
    else:
        st.dataframe(log, hide_index=True, width="stretch")
        aliases = session.alias_table()
        st.markdown("##### Alias table")
        st.caption("Every original spelling and its final name. Save it and apply it to the next case.")
        st.dataframe(aliases, hide_index=True, width="stretch")
        st.download_button(
            "Download alias table (.csv)",
            aliases.to_csv(index=False).encode(),
            file_name="alias_table.csv",
            mime="text/csv",
        )

# ----------------------------------------------------------------------
# Preview
# ----------------------------------------------------------------------
with tab_preview:
    current = session.current_sheets()
    sheet = st.selectbox("Sheet", list(current))
    df = current[sheet]
    only_changed = st.checkbox("Show only rows with normalized names")
    fields = [f for f in session.fields if f.sheet == sheet]
    if only_changed and fields:
        mask = pd.Series(False, index=df.index)
        for f in fields:
            mask |= df[f.column].fillna("").astype(str) != df[f.raw_column].fillna("").astype(str)
        df = df[mask]
    st.dataframe(df, hide_index=True, width="stretch", height=500)
    st.caption(f"{len(df)} row(s). Raw snapshot columns are kept at the end of each sheet.")
