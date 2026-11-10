import json
from pathlib import Path
import tempfile
import unittest

from signal_json_import import load_signal_json, inspect_media_paths


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
