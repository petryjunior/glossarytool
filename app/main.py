from __future__ import annotations

import csv
import os
import re
import uuid
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from threading import Lock
from typing import Dict, List, Optional

import chardet
import pandas as pd
from fastapi import Depends, FastAPI, File, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

SESSION_COOKIE = "glossary_session_id"
SESSION_STORE: Dict[str, "SessionState"] = {}
SESSION_LOCK = Lock()


LARGE_GLOSSARY_LIMIT = 10000


@dataclass
class Glossary:
    id: str
    filename: str
    display_name: str
    dataframe: pd.DataFrame
    term_column: str
    selected: bool = True
    preload_terms: bool = True


@dataclass
class SessionState:
    glossaries: List[Glossary] = field(default_factory=list)


BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR.parent / "frontend"

app = FastAPI(title="Glossary Lookup Tool")
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


def detect_csv_delimiter(content: bytes) -> str:
    sample = content[:4096]
    encoding = chardet.detect(sample).get("encoding") or "utf-8"
    try:
        decoded_sample = sample.decode(encoding, errors="ignore")
    except (LookupError, UnicodeDecodeError):
        decoded_sample = sample.decode("utf-8", errors="ignore")

    try:
        dialect = csv.Sniffer().sniff(decoded_sample, delimiters=[",", ";", "\t"])
        return dialect.delimiter
    except csv.Error:
        return ","


def get_session_state(request: Request, response: Response) -> SessionState:
    session_id = request.cookies.get(SESSION_COOKIE)
    with SESSION_LOCK:
        if not session_id or session_id not in SESSION_STORE:
            session_id = str(uuid.uuid4())
            SESSION_STORE[session_id] = SessionState()
        response.set_cookie(
            key=SESSION_COOKIE,
            value=session_id,
            httponly=True,
            samesite="lax",
            max_age=60 * 60 * 24,
        )
        return SESSION_STORE[session_id]


def summarize_glossary(glossary: Glossary) -> Dict[str, object]:
    return {
        "id": glossary.id,
        "name": glossary.display_name,
        "selected": glossary.selected,
        "terms_count": int(glossary.dataframe.shape[0]),
        "preload_terms": glossary.preload_terms,
    }


def normalize_text(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value)
    return text.replace("_x000D_", "\n").strip()


def filter_series(series: pd.Series, term: str, exact: bool, whole_word: bool) -> pd.Series:
    clean = series.dropna().astype(str)
    if not term:
        return clean
    trimmed = term.strip()
    if not trimmed:
        return clean
    term_lower = trimmed.lower()
    if exact:
        return clean[clean.str.lower() == term_lower]
    if whole_word:
        pattern = fr"\b{re.escape(trimmed)}\b"
        return clean[clean.str.contains(pattern, case=False, regex=True)]
    return clean[clean.str.contains(re.escape(trimmed), case=False, regex=True)]


def build_term_list(
    session: SessionState, search: str, exact: bool, whole_word: bool
) -> List[str]:
    terms = set()
    sanitized_search = search.strip()
    include_large = bool(sanitized_search)
    for glossary in session.glossaries:
        if not glossary.selected:
            continue
        if not include_large and not glossary.preload_terms:
            continue
        term_series = glossary.dataframe[glossary.term_column]
        matched = filter_series(term_series, sanitized_search if include_large else "", exact, whole_word)
        terms.update({value.strip() for value in matched.tolist() if value.strip()})
    if not search:
        return sorted(terms, key=str.casefold)
    return sorted(terms, key=str.casefold)


def find_glossary(session: SessionState, glossary_id: str) -> Glossary:
    for glossary in session.glossaries:
        if glossary.id == glossary_id:
            return glossary
    raise HTTPException(status_code=404, detail="Glossary not found")


@app.get("/", response_class=FileResponse)
def serve_index() -> FileResponse:
    index_path = FRONTEND_DIR / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=404, detail="Frontend not built yet")
    return FileResponse(index_path)


@app.get("/api/glossaries")
def list_glossaries(session: SessionState = Depends(get_session_state)) -> Dict[str, List[Dict[str, object]]]:
    return {"glossaries": [summarize_glossary(g) for g in session.glossaries]}


@app.post("/api/glossaries/upload")
async def upload_glossaries(
    files: List[UploadFile] = File(...),
    session: SessionState = Depends(get_session_state),
) -> Dict[str, List[Dict[str, object]]]:
    if not files:
        raise HTTPException(status_code=400, detail="Please provide one or more glossary files.")

    new_glossaries = []
    for upload in files:
        suffix = Path(upload.filename).suffix.lower()
        content = await upload.read()
        if not content:
            continue
        try:
            if suffix == ".csv":
                delimiter = detect_csv_delimiter(content)
                encoding = chardet.detect(content).get("encoding") or "utf-8"
                df = pd.read_csv(BytesIO(content), delimiter=delimiter, encoding=encoding, engine="python")
            elif suffix in {".xlsx", ".xls"}:
                df = pd.read_excel(BytesIO(content), engine="openpyxl")
            else:
                continue
        except (pd.errors.ParserError, ValueError):
            continue

        if df.empty:
            continue

        glossary = Glossary(
            id=str(uuid.uuid4()),
            filename=upload.filename,
            display_name=os.path.basename(upload.filename),
            dataframe=df,
            term_column=str(df.columns[0]),
            selected=True,
            preload_terms=len(df) <= LARGE_GLOSSARY_LIMIT,
        )
        session.glossaries.append(glossary)
        new_glossaries.append(glossary)

    if not new_glossaries:
        raise HTTPException(status_code=400, detail="No glossaries could be parsed.")

    return {"glossaries": [summarize_glossary(g) for g in session.glossaries]}


@app.post("/api/glossaries/{glossary_id}/selection")
def update_selection(
    glossary_id: str,
    payload: Dict[str, bool],
    session: SessionState = Depends(get_session_state),
) -> Dict[str, object]:
    glossary = find_glossary(session, glossary_id)
    glossary.selected = bool(payload.get("selected", False))
    return {"id": glossary.id, "selected": glossary.selected}


@app.get("/api/terms")
def search_terms(
    search: str = "",
    exact: bool = False,
    whole_word: bool = False,
    session: SessionState = Depends(get_session_state),
) -> Dict[str, List[str]]:
    terms = build_term_list(session, search, exact, whole_word)
    return {"terms": terms}


@app.get("/api/terms/details/{term}")
def term_details(
    term: str,
    session: SessionState = Depends(get_session_state),
) -> Dict[str, object]:
    term_lower = term.lower()
    results = []
    for glossary in session.glossaries:
        if not glossary.selected:
            continue
        series = glossary.dataframe[glossary.term_column]
        mask = series.astype(str, errors="ignore").str.lower() == term_lower
        matches = glossary.dataframe.loc[mask]
        if matches.empty:
            continue
        rows = []
        for _, row in matches.iterrows():
            row_data = {
                column: normalize_text(row[column])
                for column in glossary.dataframe.columns
            }
            rows.append(row_data)
        results.append({"glossary": glossary.display_name, "rows": rows})
    return {"term": term, "results": results}

