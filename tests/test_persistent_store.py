"""Test persistent_store bằng backend giả lập (không gọi mạng). Chạy: python tests/test_persistent_store.py"""
import sqlite3
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import persistent_store as ps


class FakeBackend:
    branch = "data-store"

    def __init__(self):
        self.blobs: dict[str, bytes] = {}
        self.files: dict[str, str] = {}

    def ensure_branch(self): pass
    def list_files(self): return dict(self.files)
    def get_blob(self, sha): return self.blobs[sha]

    def put_file(self, path, data, prev_sha, message):
        if self.files.get(path) != prev_sha and path in self.files:
            raise ps.ConflictError(path)
        sha = ps.git_blob_sha(data)
        self.blobs[sha] = data
        self.files[path] = sha
        return sha


def make_master(path: Path, codes):
    df = pd.DataFrame({"Mã hàng": codes, "Tên kho": "K", "Tên hàng": [f"v{c}" for c in codes], "ĐVT": "Cái"})
    df.to_excel(path, index=False)


def make_db(path: Path, rows):
    conn = sqlite3.connect(path)
    conn.execute("CREATE TABLE IF NOT EXISTS approved_materials (code TEXT PRIMARY KEY, name TEXT)")
    conn.execute("CREATE TABLE IF NOT EXISTS alias_mapping (id INTEGER PRIMARY KEY AUTOINCREMENT, raw_name TEXT UNIQUE, target_code TEXT)")
    for c in rows:
        conn.execute("INSERT OR REPLACE INTO approved_materials VALUES (?,?)", (c, "n" + c))
        conn.execute("INSERT OR IGNORE INTO alias_mapping (raw_name, target_code) VALUES (?,?)", ("raw" + c, c))
    conn.commit(); conn.close()


def db_codes(path: Path):
    conn = sqlite3.connect(path)
    r = sorted(x[0] for x in conn.execute("SELECT code FROM approved_materials"))
    a = sorted(x[0] for x in conn.execute("SELECT raw_name FROM alias_mapping"))
    conn.close()
    return r, a


def master_codes(path: Path):
    return sorted(pd.read_excel(path, dtype={"Mã hàng": str})["Mã hàng"].astype(str))


def setup_instance(root: Path, backend, master_codes_list, db_rows):
    (root / "Master_Inventory").mkdir(parents=True, exist_ok=True)
    make_master(root / "Master_Inventory" / "Master_Inventory.xlsx", master_codes_list)
    make_db(root / "materials_memory.db", db_rows)
    return ps.PersistentStore(root, backend)


def targets(root):
    return ({root / "Master_Inventory" / "Master_Inventory.xlsx": ps.KIND_MASTER,
             root / "materials_memory.db": ps.KIND_SQLITE},
            {root / "Processed_Reports": ps.KIND_IMMUTABLE})


def main():
    backend = FakeBackend()
    with tempfile.TemporaryDirectory() as t:
        t = Path(t)

        # --- Cloud lần 1: dữ liệu từ git, chưa có gì trên remote → đẩy lên
        c1 = t / "cloud1"
        s1 = setup_instance(c1, backend, ["01.0001", "01.0002"], ["01.0001"])
        f, d = targets(c1); s1.startup_sync(f, d)
        assert "Master_Inventory/Master_Inventory.xlsx" in backend.files, "master chưa được đẩy lên"

        # --- Người dùng duyệt thêm mã mới trên Cloud → persist
        make_master(c1 / "Master_Inventory" / "Master_Inventory.xlsx", ["01.0001", "01.0002", "03.9999"])
        make_db(c1 / "materials_memory.db", ["03.9999"])
        (c1 / "Processed_Reports" / "2026").mkdir(parents=True)
        (c1 / "Processed_Reports" / "2026" / "bc1.xlsx").write_bytes(b"report-1")
        s1.persist(c1 / "Master_Inventory" / "Master_Inventory.xlsx", ps.KIND_MASTER)
        s1.persist(c1 / "materials_memory.db", ps.KIND_SQLITE)
        s1.persist(c1 / "Processed_Reports" / "2026" / "bc1.xlsx", ps.KIND_IMMUTABLE)
        assert s1.flush(), "worker không xong"
        assert "Processed_Reports/2026/bc1.xlsx" in backend.files

        # --- Cloud khởi động lại: đĩa bị xóa, git cấp lại bản gốc cũ → phải lấy lại dữ liệu
        c2 = t / "cloud2"
        s2 = setup_instance(c2, backend, ["01.0001", "01.0002"], ["01.0001"])
        f, d = targets(c2); s2.startup_sync(f, d)
        assert "03.9999" in master_codes(c2 / "Master_Inventory" / "Master_Inventory.xlsx"), "mất mã đã duyệt trên master"
        assert "03.9999" in db_codes(c2 / "materials_memory.db")[0], "mất mã đã duyệt trong DB"
        assert (c2 / "Processed_Reports" / "2026" / "bc1.xlsx").read_bytes() == b"report-1", "mất báo cáo"

        # --- Cả hai phía cùng đổi → phải hợp nhất, không mất dòng nào
        make_master(c2 / "Master_Inventory" / "Master_Inventory.xlsx", ["01.0001", "01.0002", "03.9999", "03.A"])
        make_db(c2 / "materials_memory.db", ["03.A"])
        s2.persist(c2 / "Master_Inventory" / "Master_Inventory.xlsx", ps.KIND_MASTER)
        s2.persist(c2 / "materials_memory.db", ps.KIND_SQLITE)
        assert s2.flush()

        c3 = t / "local3"   # máy khác: có thay đổi riêng chưa đồng bộ
        s3 = setup_instance(c3, backend, ["01.0001", "01.0002", "03.B"], ["03.B"])
        f, d = targets(c3); s3.startup_sync(f, d)
        m = master_codes(c3 / "Master_Inventory" / "Master_Inventory.xlsx")
        for need in ["03.B", "03.9999", "03.A", "01.0001"]:
            assert need in m, f"hợp nhất master mất {need}: {m}"
        r, a = db_codes(c3 / "materials_memory.db")
        for need in ["03.B", "03.9999", "03.A"]:
            assert need in r, f"hợp nhất DB mất {need}: {r}"
        assert len(m) == len(set(m)), "hợp nhất master bị trùng mã"
        assert (c3 / "backup" / "persist").exists(), "không có bản sao lưu trước khi hợp nhất"

        # --- Lỗi mạng: app vẫn chạy bằng dữ liệu local
        class Broken(FakeBackend):
            def ensure_branch(self): raise RuntimeError("mất mạng")
        c4 = t / "local4"
        s4 = setup_instance(c4, Broken(), ["01.0001"], ["01.0001"])
        f, d = targets(c4); s4.startup_sync(f, d)
        assert "LỖI" in s4.status and master_codes(c4 / "Master_Inventory" / "Master_Inventory.xlsx") == ["01.0001"]

        # --- Không có token → tắt hẳn
        s5 = ps.PersistentStore(t / "x", None)
        s5.persist(t / "x" / "a", ps.KIND_IMMUTABLE)
        assert not s5.enabled

    print("TẤT CẢ TEST PERSISTENT STORE ĐẠT")


if __name__ == "__main__":
    main()
