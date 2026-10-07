import os
import re
import sqlite3
from difflib import SequenceMatcher, get_close_matches

from flask import Flask, render_template, request, abort, g

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Stats_lang.db")
PER_PAGE = 50

WORD_RE = re.compile(r"\w+", re.UNICODE)

app = Flask(__name__)


def tokenize(text):
    return WORD_RE.findall(text.lower())


def fuzzy_threshold(token_len):
    # shorter tokens need a tighter match to avoid noisy false positives
    if token_len <= 3:
        return 0.90
    if token_len <= 5:
        return 0.80
    return 0.72


def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()


def fuzzy_search(rows, columns, q):
    """Filter+rank rows against a free-text query with typo tolerance.

    Each query word must match somewhere in the row, either as an exact
    substring (as before) or, failing that, as a fuzzy match against one of
    the row's words. Returns (matched_rows, suggestions) where suggestions
    maps a query word with no exact hit anywhere in the table to the closest
    real word found, for a "did you mean" hint.
    """
    query_tokens = tokenize(q)
    if not query_tokens:
        for row in rows:
            row["_score"] = 1.0
        return rows, {}

    row_texts = []
    row_tokens_list = []
    vocabulary = set()
    for row in rows:
        text = " ".join(str(row[c]) for c in columns if row[c] is not None).lower()
        row_texts.append(text)
        rtokens = tokenize(text)
        row_tokens_list.append(rtokens)
        vocabulary.update(rtokens)

    suggestions = {}
    for qt in query_tokens:
        if any(qt in rt for rt in row_texts):
            continue
        close = get_close_matches(qt, vocabulary, n=1, cutoff=0.6)
        if close and close[0] != qt:
            suggestions[qt] = close[0]

    matched = []
    for row, text, rtokens in zip(rows, row_texts, row_tokens_list):
        token_scores = []
        ok = True
        for qt in query_tokens:
            if qt in text:
                token_scores.append(1.0)
                continue
            thresh = fuzzy_threshold(len(qt))
            best = max((similarity(qt, rt) for rt in rtokens), default=0.0)
            if best >= thresh:
                token_scores.append(best)
            else:
                ok = False
                break
        if ok:
            row["_score"] = sum(token_scores) / len(token_scores)
            matched.append(row)

    return matched, suggestions


def corrected_query(q, suggestions):
    corrected = q
    for typo, fix in suggestions.items():
        corrected = re.sub(r"(?i)\b" + re.escape(typo) + r"\b", fix, corrected)
    return corrected


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def get_tables():
    db = get_db()
    rows = db.execute(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    ).fetchall()
    return [r["name"] for r in rows]


def get_columns(table):
    db = get_db()
    rows = db.execute(f'PRAGMA table_info("{table}")').fetchall()
    return [r["name"] for r in rows]


def prettify(name):
    return name.replace("_lang", "").replace("_", " ").strip()


def table_meta():
    db = get_db()
    meta = []
    for t in get_tables():
        count = db.execute(f'SELECT COUNT(*) AS n FROM "{t}"').fetchone()["n"]
        cols = get_columns(t)
        meta.append({
            "name": t,
            "label": prettify(t),
            "rows": count,
            "cols": len(cols),
        })
    return meta


@app.route("/")
def index():
    meta = table_meta()
    total_rows = sum(m["rows"] for m in meta)
    return render_template(
        "index.html",
        tables=meta,
        total_rows=total_rows,
        total_tables=len(meta),
        active_table=None,
    )


@app.route("/table/<table>")
def table_view(table):
    tables = get_tables()
    if table not in tables:
        abort(404)

    db = get_db()
    columns = get_columns(table)

    q = request.args.get("q", "").strip()
    sort = request.args.get("sort", "")
    direction = request.args.get("dir", "asc")
    if direction not in ("asc", "desc"):
        direction = "asc"
    page = request.args.get("page", 1, type=int)
    if page < 1:
        page = 1

    filter_col_candidates = [c for c in ("Jahr", "Kennzahl", "Wert-Typ") if c in columns]
    active_filters = {}
    filter_options = {}
    where_clauses = []
    params = []

    for fc in filter_col_candidates:
        opts = [
            r[0] for r in db.execute(
                f'SELECT DISTINCT "{fc}" FROM "{table}" WHERE "{fc}" IS NOT NULL ORDER BY "{fc}"'
            ).fetchall()
        ]
        filter_options[fc] = opts
        val = request.args.get(fc, "").strip()
        if val:
            active_filters[fc] = val
            where_clauses.append(f'"{fc}" = ?')
            params.append(val)

    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

    # Dropdown filters run in SQL (cheap, exact). Free-text search runs in
    # Python below so it can score fuzzy/typo matches, not just substrings —
    # table sizes here (max ~2200 rows) make that trivially fast.
    all_rows = [dict(r) for r in db.execute(
        f'SELECT rowid AS _rowid, * FROM "{table}" {where_sql}', params
    ).fetchall()]

    suggestions = {}
    used_fuzzy = False
    if q:
        all_rows, suggestions = fuzzy_search(all_rows, columns, q)
        used_fuzzy = any(row["_score"] < 1.0 for row in all_rows)

    if sort in columns:
        non_null = [r for r in all_rows if r[sort] is not None]
        null_rows = [r for r in all_rows if r[sort] is None]

        def key(row):
            v = row[sort]
            try:
                return (0, float(v))
            except (TypeError, ValueError):
                return (1, str(v).lower())

        non_null.sort(key=key, reverse=(direction == "desc"))
        all_rows = non_null + null_rows
    elif q:
        all_rows.sort(key=lambda r: r["_score"], reverse=True)

    total = len(all_rows)
    total_pages = max(1, (total + PER_PAGE - 1) // PER_PAGE)
    page = min(page, total_pages)
    offset = (page - 1) * PER_PAGE
    rows = all_rows[offset:offset + PER_PAGE]

    suggestion_query = corrected_query(q, suggestions) if suggestions else None
    if suggestion_query == q:
        suggestion_query = None

    return render_template(
        "table.html",
        table=table,
        label=prettify(table),
        columns=columns,
        rows=rows,
        total=total,
        page=page,
        total_pages=total_pages,
        q=q,
        sort=sort,
        direction=direction,
        filter_col_candidates=filter_col_candidates,
        filter_options=filter_options,
        active_filters=active_filters,
        tables=table_meta(),
        active_table=table,
        used_fuzzy=used_fuzzy,
        suggestion_query=suggestion_query,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
