"""
LƯU TRỮ BỀN VỮNG DỮ LIỆU LÊN NHÁNH GITHUB RIÊNG (data-store)

Vấn đề: Streamlit Cloud xóa sạch mọi file ghi xuống đĩa mỗi khi app khởi động lại
(redeploy, ngủ đông, cập nhật code). Mã đã duyệt, Master_Inventory và báo cáo bị mất.

Giải pháp: mỗi file dữ liệu quan trọng được đẩy lên nhánh `data-store` của repo GitHub
(Streamlit Cloud chỉ theo dõi nhánh main nên KHÔNG bị redeploy khi lưu dữ liệu).
Khi app khởi động, dữ liệu được kéo về và hợp nhất an toàn với bản đang có.

Cấu hình (secrets_config.get_secret):
  GITHUB_TOKEN       — fine-grained token, quyền Contents: Read and write cho repo này
  GITHUB_REPO        — "owner/repo" (mặc định lấy từ PERSIST_DEFAULT_REPO)
  GITHUB_DATA_BRANCH — mặc định "data-store"
Không có GITHUB_TOKEN → module tự tắt, app chạy y như cũ (chỉ lưu local).

Quy tắc đồng bộ mỗi file (so sánh git-blob-sha):
  base = sha lần đồng bộ trước (manifest local) hoặc sha bản phát hành main mà Cloud đã thấy
  local == remote        → không làm gì
  local == base          → remote mới hơn → tải về
  remote == base         → local mới hơn → đẩy lên
  cả hai cùng đổi        → HỢP NHẤT (Master: gộp theo Mã hàng; DB: gộp theo khóa; báo cáo: hợp)
Không bao giờ ghi đè mất dữ liệu: bản local bị thay luôn được sao lưu *.bak-<thời gian>.
"""

from __future__ import annotations

import base64
import hashlib
import json
import queue
import shutil
import sqlite3
import threading
import time
from pathlib import Path

from secrets_config import get_secret

PERSIST_DEFAULT_REPO = "phongktcfc-tech/quanlyvattukythuatCFC"
REMOTE_MANIFEST_PATH = "_persist/manifest.json"
LOCAL_MANIFEST_NAME = ".persist_manifest.json"
HTTP_TIMEOUT = 15

KIND_MASTER = "master_excel"
KIND_SQLITE = "sqlite"
KIND_IMMUTABLE = "immutable"  # báo cáo: tên file duy nhất, không bao giờ sửa


def git_blob_sha(data: bytes) -> str:
    return hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()


def file_sha(path: Path) -> str | None:
    try:
        return git_blob_sha(path.read_bytes()) if path.is_file() else None
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════════════════
# GITHUB API CLIENT (tối thiểu, chỉ dùng requests)
# ═══════════════════════════════════════════════════════════════════════════
class GitHubBackend:
    def __init__(self, token: str, repo: str, branch: str):
        import requests
        self._rq = requests
        self.repo = repo
        self.branch = branch
        self.api = f"https://api.github.com/repos/{repo}"
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }

    def _req(self, method: str, url: str, **kw):
        return self._rq.request(method, url, headers=self.headers, timeout=HTTP_TIMEOUT, **kw)

    def ensure_branch(self) -> None:
        r = self._req("GET", f"{self.api}/git/ref/heads/{self.branch}")
        if r.status_code == 200:
            return
        if r.status_code != 404:
            r.raise_for_status()
        # Tạo nhánh mồ côi (orphan) chỉ chứa README, không kéo theo code
        readme = "Nhánh lưu dữ liệu tự động của ứng dụng. KHÔNG merge vào main.\n"
        blob = self._req("POST", f"{self.api}/git/blobs", json={"content": readme, "encoding": "utf-8"})
        blob.raise_for_status()
        tree = self._req("POST", f"{self.api}/git/trees", json={"tree": [
            {"path": "README.md", "mode": "100644", "type": "blob", "sha": blob.json()["sha"]}]})
        tree.raise_for_status()
        commit = self._req("POST", f"{self.api}/git/commits", json={
            "message": "init data-store", "tree": tree.json()["sha"], "parents": []})
        commit.raise_for_status()
        ref = self._req("POST", f"{self.api}/git/refs", json={
            "ref": f"refs/heads/{self.branch}", "sha": commit.json()["sha"]})
        if ref.status_code not in (201, 422):  # 422: tiến trình khác vừa tạo
            ref.raise_for_status()

    def list_files(self) -> dict[str, str]:
        """{đường_dẫn: blob_sha} của toàn bộ file trên nhánh."""
        r = self._req("GET", f"{self.api}/git/trees/{self.branch}", params={"recursive": "1"})
        r.raise_for_status()
        return {e["path"]: e["sha"] for e in r.json().get("tree", []) if e.get("type") == "blob"}

    def get_blob(self, sha: str) -> bytes:
        r = self._req("GET", f"{self.api}/git/blobs/{sha}")
        r.raise_for_status()
        return base64.b64decode(r.json()["content"])

    def put_file(self, path: str, data: bytes, prev_sha: str | None, message: str) -> str:
        """Ghi file; trả về blob sha mới. Ném ConflictError nếu prev_sha đã cũ."""
        body = {"message": message, "content": base64.b64encode(data).decode("ascii"), "branch": self.branch}
        if prev_sha:
            body["sha"] = prev_sha
        r = self._req("PUT", f"{self.api}/contents/{path}", json=body)
        if r.status_code in (409, 422):
            raise ConflictError(path)
        r.raise_for_status()
        return r.json()["content"]["sha"]


class ConflictError(Exception):
    pass


# ═══════════════════════════════════════════════════════════════════════════
# HỢP NHẤT DỮ LIỆU KHI CẢ HAI PHÍA CÙNG THAY ĐỔI
# ═══════════════════════════════════════════════════════════════════════════
def merge_master_excel(local_bytes: bytes, remote_bytes: bytes) -> bytes:
    """Giữ nguyên mọi dòng local, bổ sung các Mã hàng chỉ có ở remote (mã mới được duyệt trên Cloud)."""
    import io
    import pandas as pd
    dt = {"Mã hàng": str, "Tên kho": str, "Tên hàng": str, "ĐVT": str}
    df_l = pd.read_excel(io.BytesIO(local_bytes), dtype=dt)
    df_r = pd.read_excel(io.BytesIO(remote_bytes), dtype=dt)
    if "Mã hàng" not in df_l.columns or "Mã hàng" not in df_r.columns:
        return local_bytes
    codes_l = set(df_l["Mã hàng"].dropna().astype(str).str.strip())
    extra = df_r[~df_r["Mã hàng"].fillna("").astype(str).str.strip().isin(codes_l)]
    extra = extra[extra["Mã hàng"].fillna("").astype(str).str.strip() != ""]
    if extra.empty:
        return local_bytes
    merged = pd.concat([df_l, extra[[c for c in extra.columns if c in df_l.columns]]], ignore_index=True)
    merged = merged.sort_values(by="Mã hàng").reset_index(drop=True)
    buf = io.BytesIO()
    merged.to_excel(buf, index=False)
    return buf.getvalue()


def merge_sqlite(local_path: Path, remote_bytes: bytes) -> None:
    """Chép vào DB local mọi dòng remote chưa có (theo khóa chính / UNIQUE). Local thắng khi trùng khóa."""
    tmp = local_path.with_suffix(".remote.tmp")
    tmp.write_bytes(remote_bytes)
    try:
        conn = sqlite3.connect(local_path)
        conn.execute("ATTACH DATABASE ? AS r", (str(tmp),))
        tables = [row[0] for row in conn.execute(
            "SELECT name FROM r.sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
        for t in tables:
            exists = conn.execute("SELECT 1 FROM main.sqlite_master WHERE type='table' AND name=?", (t,)).fetchone()
            if not exists:
                ddl = conn.execute("SELECT sql FROM r.sqlite_master WHERE type='table' AND name=?", (t,)).fetchone()[0]
                conn.execute(ddl)
            cols_main = [c[1] for c in conn.execute(f'PRAGMA main.table_info("{t}")')]
            cols_r = [c[1] for c in conn.execute(f'PRAGMA r.table_info("{t}")')]
            # Bỏ cột id tự tăng để không đè id; khóa UNIQUE/PRIMARY KEY khác sẽ chặn trùng
            pk_auto = [c[1] for c in conn.execute(f'PRAGMA main.table_info("{t}")') if c[5] == 1 and c[2].upper() == "INTEGER"]
            cols = [c for c in cols_r if c in cols_main and c not in pk_auto]
            if not cols:
                continue
            col_sql = ", ".join(f'"{c}"' for c in cols)
            conn.execute(f'INSERT OR IGNORE INTO main."{t}" ({col_sql}) SELECT {col_sql} FROM r."{t}"')
        conn.commit()
        conn.execute("DETACH DATABASE r")
        conn.close()
    finally:
        try:
            tmp.unlink()
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════════
# BỘ ĐỒNG BỘ
# ═══════════════════════════════════════════════════════════════════════════
class PersistentStore:
    def __init__(self, base_dir: Path, backend=None):
        self.base_dir = Path(base_dir)
        self.backend = backend
        self.enabled = backend is not None
        self.status = "TẮT (chưa cấu hình GITHUB_TOKEN)" if not self.enabled else "đang khởi tạo"
        self.last_error = ""
        self._lock = threading.RLock()
        self._remote: dict[str, str] = {}
        self._remote_manifest: dict = {}
        self._manifest_path = self.base_dir / LOCAL_MANIFEST_NAME
        self._local_manifest: dict = self._read_local_manifest()
        self._kinds: dict[str, str] = {}
        self._queue: "queue.Queue[str]" = queue.Queue()
        self._worker: threading.Thread | None = None
        self.synced = False

    # ── manifest ──
    def _read_local_manifest(self) -> dict:
        try:
            return json.loads(self._manifest_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _write_local_manifest(self) -> None:
        try:
            self._manifest_path.write_text(json.dumps(self._local_manifest, indent=1, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            print(f"[persist] Không ghi được manifest local: {e}")

    def _rel(self, path: Path) -> str:
        return Path(path).resolve().relative_to(self.base_dir.resolve()).as_posix()

    def _backup_local(self, path: Path) -> None:
        if path.is_file():
            stamp = time.strftime("%Y%m%d_%H%M%S")
            bak_dir = self.base_dir / "backup" / "persist"
            bak_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, bak_dir / f"{path.name}.bak-{stamp}")

    # ── đồng bộ khởi động ──
    def startup_sync(self, files: dict[Path, str], dirs: dict[Path, str]) -> None:
        """files: {đường_dẫn_file: KIND}; dirs: {thư_mục: KIND} (đồng bộ mọi file bên trong)."""
        if not self.enabled or self.synced:
            return
        with self._lock:
            try:
                self.backend.ensure_branch()
                self._remote = self.backend.list_files()
                if REMOTE_MANIFEST_PATH in self._remote:
                    try:
                        self._remote_manifest = json.loads(self.backend.get_blob(self._remote[REMOTE_MANIFEST_PATH]))
                    except Exception:
                        self._remote_manifest = {}
                first_local_run = not self._local_manifest
                manifest_changed = False

                targets: dict[str, str] = {self._rel(p): k for p, k in files.items()}
                for d, kind in dirs.items():
                    d = Path(d)
                    d.mkdir(parents=True, exist_ok=True)
                    rel_d = self._rel(d)
                    for p in d.rglob("*"):
                        if p.is_file() and p.name != ".gitkeep":
                            targets[self._rel(p)] = kind
                    for rp in self._remote:
                        if rp.startswith(rel_d + "/") and not rp.endswith("/.gitkeep"):
                            targets[rp] = kind
                self._kinds.update(targets)

                for rel, kind in targets.items():
                    shipped_sha = file_sha(self.base_dir / rel)
                    self._sync_one(rel, kind)
                    if first_local_run and shipped_sha and kind != KIND_IMMUTABLE:
                        if self._remote_manifest.get(rel) != shipped_sha:
                            self._remote_manifest[rel] = shipped_sha
                            manifest_changed = True
                if manifest_changed:
                    self._push_remote_manifest()
                self._write_local_manifest()
                self.status = f"BẬT — nhánh {self.backend.branch}"
            except Exception as e:
                self.last_error = str(e)
                self.status = f"LỖI đồng bộ ({type(e).__name__}) — đang dùng dữ liệu local"
                print(f"[persist] startup_sync lỗi: {e}")
            finally:
                self.synced = True
        self._start_worker()

    def _sync_one(self, rel: str, kind: str) -> None:
        path = self.base_dir / rel
        local = file_sha(path)
        remote = self._remote.get(rel)
        base = self._local_manifest.get(rel) or self._remote_manifest.get(rel)

        if remote is None and local is None:
            return
        if remote == local:
            self._local_manifest[rel] = local
            return
        if remote is None:
            self._upload(rel)
            return
        if local is None or kind == KIND_IMMUTABLE:
            if local is not None:
                return  # báo cáo trùng tên khác nội dung: giữ bản local
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(self.backend.get_blob(remote))
            self._local_manifest[rel] = remote
            return
        if local == base:  # local không đổi từ lần đồng bộ trước → remote mới hơn
            self._backup_local(path)
            path.write_bytes(self.backend.get_blob(remote))
            self._local_manifest[rel] = remote
            return
        if remote == base:  # remote không đổi → local mới hơn
            self._upload(rel)
            return
        # Cả hai cùng đổi → hợp nhất rồi đẩy lên
        self._backup_local(path)
        remote_bytes = self.backend.get_blob(remote)
        if kind == KIND_MASTER:
            path.write_bytes(merge_master_excel(path.read_bytes(), remote_bytes))
        elif kind == KIND_SQLITE:
            merge_sqlite(path, remote_bytes)
        self._upload(rel)

    def _upload(self, rel: str) -> None:
        path = self.base_dir / rel
        if not path.is_file():
            return
        data = path.read_bytes()
        sha = git_blob_sha(data)
        if self._remote.get(rel) == sha:
            self._local_manifest[rel] = sha
            return
        try:
            new_sha = self.backend.put_file(rel, data, self._remote.get(rel), f"data: cập nhật {rel}")
        except ConflictError:
            # Tiến trình khác vừa ghi: làm mới danh sách và hợp nhất lại một lần
            self._remote = self.backend.list_files()
            kind = self._kinds.get(rel, KIND_IMMUTABLE)
            remote = self._remote.get(rel)
            if remote and remote != sha and kind in (KIND_MASTER, KIND_SQLITE):
                remote_bytes = self.backend.get_blob(remote)
                if kind == KIND_MASTER:
                    path.write_bytes(merge_master_excel(path.read_bytes(), remote_bytes))
                else:
                    merge_sqlite(path, remote_bytes)
                data = path.read_bytes()
                sha = git_blob_sha(data)
            new_sha = self.backend.put_file(rel, data, self._remote.get(rel), f"data: cập nhật {rel}")
        self._remote[rel] = new_sha
        self._local_manifest[rel] = new_sha

    def _push_remote_manifest(self) -> None:
        data = json.dumps(self._remote_manifest, indent=1, ensure_ascii=False).encode("utf-8")
        try:
            self._remote[REMOTE_MANIFEST_PATH] = self.backend.put_file(
                REMOTE_MANIFEST_PATH, data, self._remote.get(REMOTE_MANIFEST_PATH), "data: manifest")
        except Exception as e:
            print(f"[persist] Không ghi được manifest remote: {e}")

    # ── lưu sau mỗi lần ghi (chạy nền, không làm chậm giao diện) ──
    def persist(self, path: Path, kind: str) -> None:
        if not self.enabled:
            return
        try:
            rel = self._rel(path)
        except ValueError:
            return
        self._kinds[rel] = kind
        self._queue.put(rel)
        self._start_worker()

    def _start_worker(self) -> None:
        if not self.enabled or (self._worker and self._worker.is_alive()):
            return
        self._worker = threading.Thread(target=self._run_worker, name="persist-worker", daemon=True)
        self._worker.start()

    def _run_worker(self) -> None:
        while True:
            rel = self._queue.get()
            time.sleep(1.0)  # gom các lần ghi dồn dập
            pending = {rel}
            while not self._queue.empty():
                pending.add(self._queue.get_nowait())
            for r in pending:
                for attempt in range(3):
                    try:
                        with self._lock:
                            self._upload(r)
                            self._write_local_manifest()
                        self.last_error = ""
                        break
                    except Exception as e:
                        self.last_error = f"{r}: {e}"
                        print(f"[persist] Lỗi đẩy {r} (lần {attempt + 1}): {e}")
                        time.sleep(2 * (attempt + 1))

    def flush(self, timeout: float = 30.0) -> bool:
        """Chờ hàng đợi rỗng (dùng cho test)."""
        end = time.time() + timeout
        while time.time() < end:
            if self._queue.empty() and self._lock.acquire(timeout=0.5):
                self._lock.release()
                time.sleep(1.5)
                if self._queue.empty():
                    return True
            time.sleep(0.2)
        return False


_STORE: PersistentStore | None = None
_STORE_LOCK = threading.Lock()


def get_store(base_dir: Path) -> PersistentStore:
    """Một store duy nhất cho mỗi tiến trình (an toàn khi Streamlit chạy lại script)."""
    global _STORE
    with _STORE_LOCK:
        if _STORE is None:
            token = get_secret("GITHUB_TOKEN")
            backend = None
            if token:
                try:
                    backend = GitHubBackend(
                        token,
                        get_secret("GITHUB_REPO", PERSIST_DEFAULT_REPO),
                        get_secret("GITHUB_DATA_BRANCH", "data-store"),
                    )
                except Exception as e:
                    print(f"[persist] Không khởi tạo được GitHub backend: {e}")
            _STORE = PersistentStore(base_dir, backend)
        return _STORE
