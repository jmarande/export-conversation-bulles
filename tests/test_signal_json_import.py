import json
from pathlib import Path
import tempfile
import unittest

from signal_json_import import load_signal_json, inspect_media_paths
import pandas as pd
from docx import Document
from export_conversation_bulles import export_chat_to_docx


class SignalImportTests(unittest.TestCase):
    def test_synthetic_chat_and_media(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            payload = {
                "backup": {"snapshot": "test-synthetic"},
                "chats": [{"id": 2, "name": "Groupe exemple"}],
                "messages": [
                    {"chatId": 2, "chat": "Groupe exemple",
                     "author": {"self": True, "name": "Note to Self"},
                     "type": "standardMessage", "direction": "outgoing",
                     "dateSentIso": "2026-07-20T14:00:00Z", "body": "Bonjour",
                     "attachments": [
                         {"available": True, "extractedPath": "002-Groupe\\image01.jpg"},
                         {"available": False, "role": "quote-thumbnail", "unavailableReason": "not in backup"}]},
                    {"chatId": 2, "chat": "Groupe exemple",
                     "author": {"name": "Contact A"},
                     "type": "update", "direction": "directionless",
                     "dateSentIso": "2026-07-20T14:01:00Z",
                     "update": {"type": "groupChange", "text": "Group updated"}}
                ]
            }
            data = root / "export.json"
            data.write_text(json.dumps(payload), encoding="utf-8")
            rows, meta = load_signal_json(data)
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["From"], "Compte local (Signal)")
            self.assertEqual(rows[0]["MediaPath"], "002-Groupe/image01.jpg")
            self.assertIn("16:00:00", rows[0]["Timestamp"])
            self.assertIn("quote-thumbnail", rows[0]["Body"])
            self.assertIn("Événement Signal", rows[1]["Body"])
            self.assertEqual(meta["media_references"], 1)
            self.assertEqual(meta["unavailable_media_references"], 1)
            file = root / "medias" / "002-Groupe" / "image01.jpg"
            file.parent.mkdir(parents=True)
            file.write_bytes(b"FAKE")
            self.assertEqual(inspect_media_paths(rows, root / "medias")["found"], 1)


    def test_docx_contacts_and_chat_sections(self):
        """No evidence is lost when messages from two chats interleave in time."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            payload = {
                "recipients": [
                    {"id": 2, "type": "self", "name": "Note to Self"},
                    {"id": 10, "type": "contact", "name": "Alice", "e164": "+33600000001", "aci": "A01"},
                    {"id": 11, "type": "contact", "name": "Bob", "e164": "+33600000002", "aci": "B02"},
                    {"id": 20, "type": "group", "name": "Groupe Test", "members": [
                        {"aci": "A01", "role": "ADMINISTRATOR"},
                        {"aci": "B02", "role": "DEFAULT"}]},
                    {"id": 99, "type": "contact", "name": "Hors dossier", "e164": "+33699999999"}
                ],
                "chats": [
                    {"id": 1, "name": "Groupe Test", "type": "group", "recipientId": 20},
                    {"id": 2, "name": "Alice direct", "type": "contact", "recipientId": 10}
                ],
                "messages": [
                    {"chatId": 1, "chat": "Groupe Test", "type": "standardMessage",
                     "author": {"id": 10, "name": "Alice"},
                     "direction": "incoming", "body": "message A1", "dateSentIso": "2026-07-20T10:00:00Z"},
                    {"chatId": 2, "chat": "Alice direct", "type": "standardMessage",
                     "author": {"id": 2, "self": True},
                     "direction": "outgoing", "body": "message D1", "dateSentIso": "2026-07-20T10:01:00Z"},
                    {"chatId": 1, "chat": "Groupe Test", "type": "standardMessage",
                     "author": {"id": 11, "name": "Bob"},
                     "direction": "incoming", "body": "message A2", "dateSentIso": "2026-07-20T10:02:00Z"}
                ]
            }
            infile = root / "conversations.json"
            infile.write_text(json.dumps(payload), encoding="utf-8")
            rows, meta = load_signal_json(infile)
            self.assertEqual(len(meta["participants_by_chat"]["1"]), 3)
            self.assertTrue(all(p["name"] != "Hors dossier" for persons in meta["participants_by_chat"].values() for p in persons))
            output = root / "rapport.docx"
            export_chat_to_docx(
                df=pd.DataFrame(rows), out_path=str(output),
                col_sender="From", col_message="Body", col_direction="Direction",
                mode_side="direction", right_participant_value=None, owner_col=None,
                ts_single_col="Timestamp", ts_date_col=None, ts_time_col=None,
                display_name_map={}, empty_sender_display="Compte local",
                participant_color_map={}, default_left_color="#D7E3F1",
                default_right_color="#D8EBD5", bubble_width_cm=10,
                gutter_cm=1, bubble_spacing_pt=4, extra_cols=None,
                media_path_col="MediaPath", files_root="", search_subfolders=False,
                fuzzy_match=False, image_width_mode="fit_bubble",
                image_max_width_cm=8, compress_images=False, jpeg_quality=80,
                max_image_px=1920, keep_media_cache=False, video_thumbs=False,
                video_thumb_width_cm=8, status_col=None, status_map={},
                orientation="portrait", page_title="Restitution Signal",
                intro_text="", page_numbers=False, strict_media_lookup=True,
                signal_report=meta
            )
            self.assertTrue(output.is_file())
            document = Document(output)
            paragraphs = [p.text for p in document.paragraphs]
            self.assertTrue(any("Répertoire des contacts" in x for x in paragraphs))
            self.assertTrue(any("Conversation 1 / 2" in x for x in paragraphs))
            self.assertTrue(any("Conversation 2 / 2" in x for x in paragraphs))
            self.assertGreaterEqual(len(document.tables), 6)  # contacts, 2 metadata tables, 3 bubbles
            table_contact_text = " ".join(c.text for row in document.tables[0].rows for c in row.cells)
            self.assertIn("Alice", table_contact_text)
            self.assertIn("Bob", table_contact_text)
            self.assertNotIn("Hors dossier", table_contact_text)
            first_conversation = next(i for i,x in enumerate(paragraphs) if x.startswith("Conversation 1 / 2"))
            second_conversation = next(i for i,x in enumerate(paragraphs) if x.startswith("Conversation 2 / 2"))
            self.assertLess(first_conversation, second_conversation)
            breaks = sum(len(p._p.xpath('.//w:br[@w:type="page"]')) for p in document.paragraphs)
            self.assertEqual(breaks, 2)

    def test_reject_traversal(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "export.json"
            path.write_text(json.dumps({"chats": [], "messages": [
                {"chatId": 1, "dateSentIso": "2026-11-09T10:00:00Z",
                 "attachments": [{"available": True, "extractedPath": "..\\elsewhere\\sensitive.txt"}]}
            ]}), encoding="utf-8")
            rows, meta = load_signal_json(path)
            self.assertEqual(rows[0]["MediaPath"], "")
            self.assertIn("chemin", rows[0]["MediaStatus"])
            self.assertEqual(meta["unavailable_media_references"], 1)


if __name__ == "__main__":
    unittest.main()
