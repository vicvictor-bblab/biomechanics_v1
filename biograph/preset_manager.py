import json
import sqlite3


class PresetManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = None
        self.init_db()

    def init_db(self):
        self.conn = sqlite3.connect(self.db_path)
        cur = self.conn.cursor()
        cur.execute(
            "CREATE TABLE IF NOT EXISTS presets (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, description TEXT, tags TEXT, settings TEXT NOT NULL)"
        )
        self.conn.commit()

    def list_presets(self):
        cur = self.conn.cursor()
        cur.execute("SELECT name FROM presets ORDER BY name")
        return [r[0] for r in cur.fetchall()]

    def save_preset(self, name, description, tags, settings_dict):
        cur = self.conn.cursor()
        settings_json = json.dumps(settings_dict)
        cur.execute("SELECT id FROM presets WHERE name=?", (name,))
        existing = cur.fetchone()
        if existing:
            cur.execute(
                "UPDATE presets SET description=?, tags=?, settings=? WHERE name=?",
                (description, tags, settings_json, name),
            )
        else:
            cur.execute(
                "INSERT INTO presets (name, description, tags, settings) VALUES (?,?,?,?)",
                (name, description, tags, settings_json),
            )
        self.conn.commit()

    def load_preset(self, name):
        cur = self.conn.cursor()
        cur.execute("SELECT settings FROM presets WHERE name=?", (name,))
        row = cur.fetchone()
        if row:
            return json.loads(row[0])
        return None

    def delete_preset(self, name):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM presets WHERE name=?", (name,))
        self.conn.commit()

    def export_preset(self, name, file_path):
        cur = self.conn.cursor()
        cur.execute(
            "SELECT name, description, tags, settings FROM presets WHERE name=?",
            (name,),
        )
        row = cur.fetchone()
        if row:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "name": row[0],
                        "description": row[1],
                        "tags": row[2],
                        "settings": json.loads(row[3]),
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )

    def import_preset(self, file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.save_preset(
            data.get("name"), data.get("description"), data.get("tags"), data.get("settings", {})
        )
