# -*- coding: utf-8 -*-
"""Read-only importer for Signal V2 JSON conversation exports.

Converts an authenticated, already-decoded export into the columns expected by
Export Conversation Bulles. Does not handle encryption or change evidence.
"""
from collections import Counter
from datetime import datetime
from hashlib import sha256
from pathlib import Path, PurePosixPath
from zoneinfo import ZoneInfo
import json


def _person_name(value):
    if not isinstance(value, dict):
        return "Auteur non documenté"
    if value.get("self"):
        return "Compte local (Signal)"
    return str(value.get("name") or value.get("e164") or value.get("aci") or "Auteur non documenté")


def _fr_datetime(iso, zone):
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(str(iso).replace("Z", "+00:00"))
        return dt.astimezone(zone).strftime("%d/%m/%Y %H:%M:%S %z") if dt.tzinfo else ""
    except (ValueError, TypeError, OverflowError):
        return ""


def _relative_media_path(value):
    """Reject absolute paths, traversal and drive prefixes in untrusted exports."""
    if not isinstance(value, str) or not value.strip():
        return None
    value = value.strip().replace("\\", "/")
    if value.startswith("/") or ":" in value or "\x00" in value:
        return None
    if any(part in (".", "..") for part in value.split("/")):
        return None
    return PurePosixPath(value).as_posix()


def load_signal_json(path):
    source = Path(path)
    raw = source.read_bytes()
    data = json.loads(raw.decode("utf-8-sig"))
    if not isinstance(data, dict) or not isinstance(data.get("chats"), list) or not isinstance(data.get("messages"), list):
        raise ValueError("Export Signal JSON non reconnu : champs 'chats' et 'messages' absents.")
    try:
        paris = ZoneInfo("Europe/Paris")
    except Exception as exc:
        raise RuntimeError("Fuseau Europe/Paris indisponible ; installer tzdata.") from exc

    chats = {str(c.get("id")): str(c.get("name") or "Conversation sans nom")
             for c in data["chats"] if isinstance(c, dict)}
    rows = []
    unavailable = 0
    referenced = 0
    for index, message in enumerate(data["messages"], 1):
        if not isinstance(message, dict):
            continue
        author = message.get("author") if isinstance(message.get("author"), dict) else {}
        chat_id = str(message.get("chatId", ""))
        kind = str(message.get("type") or "")
        body = str(message.get("body") or "")
        if kind == "update":
            update = message.get("update")
            if isinstance(update, dict):
                body = "[Événement Signal] " + str(update.get("text") or update.get("type") or "Mise à jour")
        paths, warnings = [], []
        for item in message.get("attachments") or []:
            if not isinstance(item, dict):
                continue
            if item.get("available") is False:
                unavailable += 1
                warnings.append(str(item.get("role") or "média") + " indisponible : "
                                + str(item.get("unavailableReason") or "non présent"))
                continue
            rel = _relative_media_path(item.get("extractedPath"))
            if rel:
                referenced += 1
                paths.append(rel)
            else:
                unavailable += 1
                warnings.append("référence média sans chemin exploitable")
        if warnings:
            body += ("\n" if body else "") + "[Média non disponible : " + "; ".join(warnings) + "]"
        quotes = message.get("quote")
        citation = ""
        if isinstance(quotes, dict):
            citation = ("Réponse citée — " + _person_name(quotes.get("author")) + " : "
                        + str(quotes.get("body") or "") + " [" + str(quotes.get("targetSentIso") or "") + "]")
        reactions = []
        for r in message.get("reactions") or []:
            if isinstance(r, dict):
                reactions.append(str(r.get("emoji") or "") + " (" + _person_name(r.get("author")) + ")")
        statuses = []
        for s in message.get("sendStatus") or []:
            if isinstance(s, dict):
                statuses.append(_person_name(s.get("recipient")) + " : " + str(s.get("status") or ""))
        rows.append({
            "Conversation": str(message.get("chat") or chats.get(chat_id) or ("Conversation " + chat_id)),
            "ChatID": chat_id,
            "From": _person_name(author),
            "Body": body,
            "Direction": str(message.get("direction") or "directionless"),
            "Timestamp": _fr_datetime(message.get("dateSentIso"), paris),
            "TimestampUTC": str(message.get("dateSentIso") or ""),
            "ReceivedUTC": str(message.get("dateReceivedIso") or ""),
            "MediaPath": ";".join(paths),
            "MediaStatus": "; ".join(warnings),
            "Type": kind,
            "Citation": citation,
            "Reactions": "Réactions : " + "; ".join(reactions) if reactions else "",
            "SendStatus": "; ".join(statuses),
            "AuthorID": str(author.get("aci") or author.get("id") or ""),
            "SourceRow": index
        })
    rows.sort(key=lambda r: (r["TimestampUTC"], r["SourceRow"]))
    backup = data.get("backup") if isinstance(data.get("backup"), dict) else {}
    return rows, {
        "source_sha256": sha256(raw).hexdigest(),
        "snapshot": backup.get("snapshot"),
        "backup_time": backup.get("backupTimeIso"),
        "per_chat": dict(Counter(row["Conversation"] for row in rows)),
        "messages": len(rows),
        "media_references": referenced,
        "unavailable_media_references": unavailable
    }


def inspect_media_paths(rows, media_root):
    """Exact relative-path joins only; no fuzzy filename matches."""
    base = Path(media_root)
    found, missing = 0, []
    for row in rows:
        for value in str(row.get("MediaPath") or "").split(";"):
            if not value:
                continue
            rel = _relative_media_path(value)
            if rel and (base / Path(rel)).is_file():
                found += 1
            else:
                missing.append({"source_row": row.get("SourceRow"), "path": value})
    return {"found": found, "expected": found + len(missing), "missing": missing}
