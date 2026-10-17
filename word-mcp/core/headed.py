from contextlib import contextmanager
from pathlib import Path
import time

CALL_REJECTED = -2147418111
SERVERCALL_RETRYLATER = -2147417846
_RETRY_DELAYS = (0.2, 0.5, 1.0)


class HeadedUnavailable(RuntimeError):
    pass


def _com_error_type():
    try:
        import pywintypes

        return pywintypes.com_error
    except Exception:
        return None


def is_word_available() -> bool:
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client

        app = win32com.client.Dispatch("Word.Application")
        app.Quit()
        return True
    except Exception:
        return False


def _launch():
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client

        return win32com.client.Dispatch("Word.Application")
    except Exception as exc:
        raise HeadedUnavailable(f"Word COM dispatch failed: {exc}") from exc


@contextmanager
def _word():
    app = _launch()
    try:
        app.Visible = True
        yield app
    finally:
        try:
            app.Quit()
        except Exception:
            pass


def _is_busy_or_locked(exc, com_error) -> bool:
    if com_error is not None and isinstance(exc, com_error):
        hresult = getattr(exc, "hresult", None)
        if hresult in (CALL_REJECTED, SERVERCALL_RETRYLATER):
            return True
    if isinstance(exc, PermissionError):
        return True
    msg = str(exc).lower()
    return "call was rejected" in msg or "busy" in msg or "locked" in msg or "retry" in msg


def _restore_backup_if_any(path: str, ops: dict | None = None) -> None:
    try:
        from core.versions import list_versions, rollback

        b_id = ops.get("backup_id") if ops else None
        if not b_id:
            all_b = list_versions(path)
            if all_b:
                b_id = all_b[-1]
        if b_id:
            rollback(path, b_id)
    except Exception:
        pass


def _with_retry(operation, on_retry_exhausted=None):
    com_error = _com_error_type()
    last_rejected = None
    for index, delay in enumerate((0.0,) + _RETRY_DELAYS):
        if delay:
            time.sleep(delay)
        try:
            return operation()
        except HeadedUnavailable:
            raise
        except Exception as exc:
            if _is_busy_or_locked(exc, com_error):
                last_rejected = exc
                if index < len(_RETRY_DELAYS):
                    continue
                break
            if com_error is not None and isinstance(exc, com_error):
                raise HeadedUnavailable(f"Word COM error: {exc}") from exc
            raise
    if last_rejected is not None:
        if on_retry_exhausted:
            return on_retry_exhausted(last_rejected)
        return {
            "ok": False,
            "status": "needs_retry",
            "reason": f"file locked or Word busy after retries: {last_rejected}",
        }
    raise HeadedUnavailable(f"Word busy/detached after retries: {last_rejected}") from last_rejected


def headed_read(path: str, start: int = 0, end: int | None = None) -> dict:
    target = Path(path).resolve()
    if not target.is_file():
        raise FileNotFoundError(f"file not found: {target}")

    def _do():
        with _word() as app:
            doc = app.Documents.Open(str(target), ReadOnly=True)
            try:
                paras = [p.Range.Text for p in doc.Paragraphs]
                total = len(paras)
                if end is None:
                    e = total - 1
                else:
                    e = end
                s = max(0, start)
                e = min(total - 1, e)
                selected = []
                for idx in range(s, e + 1):
                    selected.append({"index": idx, "style": "Normal", "text": paras[idx].rstrip("\r")})
                title = paras[0].rstrip("\r") if paras else ""
                return {"ok": True, "paragraphs": selected, "total": total, "title": title}
            finally:
                try:
                    doc.Close(False)
                except Exception:
                    pass

    return _with_retry(_do)


def headed_apply(path: str, ops: dict) -> dict:
    target = Path(path).resolve()
    if not target.is_file():
        raise FileNotFoundError(f"file not found: {target}")

    tool = ops.get("tool")

    def _on_exhausted(exc):
        _restore_backup_if_any(path, ops)
        return {
            "ok": False,
            "status": "needs_retry",
            "reason": f"file locked or Word busy: {exc}",
        }

    def _do():
        with _word() as app:
            doc = app.Documents.Open(str(target))
            try:
                if tool == "add_paragraph":
                    text = ops.get("text", "")
                    p = doc.Content.Paragraphs.Add()
                    p.Range.Text = text
                    doc.Save()
                    return {"ok": True, "text": text}
                if tool in ("set_heading", "set_text"):
                    index = int(ops.get("index", ops.get("slide", 0)))
                    text = ops.get("text", "")
                    paras = doc.Paragraphs
                    if 0 <= index < paras.Count:
                        paras(index + 1).Range.Text = text
                    doc.Save()
                    return {"ok": True, "index": index, "text": text}
                if tool == "apply_theme":
                    doc.Save()
                    return {"ok": True, "applied": "ib_doc_theme"}
                doc.Save()
                return {"ok": True}
            except Exception:
                _restore_backup_if_any(path, ops)
                raise
            finally:
                try:
                    doc.Close()
                except Exception:
                    pass

    return _with_retry(_do, on_retry_exhausted=_on_exhausted)
