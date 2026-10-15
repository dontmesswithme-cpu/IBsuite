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


def is_powerpoint_available() -> bool:
    try:
        import win32com.client

        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:
        return False


def _launch():
    try:
        import win32com.client

        return win32com.client.Dispatch("PowerPoint.Application")
    except Exception as exc:
        raise HeadedUnavailable(f"PowerPoint COM dispatch failed: {exc}") from exc


@contextmanager
def _powerpoint():
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
                raise HeadedUnavailable(f"PowerPoint COM error: {exc}") from exc
            raise
    if last_rejected is not None:
        if on_retry_exhausted:
            return on_retry_exhausted(last_rejected)
        return {
            "ok": False,
            "status": "needs_retry",
            "reason": f"file locked or PowerPoint busy after retries: {last_rejected}",
        }
    raise HeadedUnavailable(f"PowerPoint busy/detached after retries: {last_rejected}") from last_rejected


def headed_read(path: str, slide: int = 1) -> dict:
    target = Path(path).resolve()
    if not target.is_file():
        raise FileNotFoundError(f"file not found: {target}")

    def _do():
        with _powerpoint() as app:
            pres = app.Presentations.Open(str(target), ReadOnly=True, WithWindow=True)
            try:
                if slide < 1 or slide > pres.Slides.Count:
                    raise ValueError(f"unknown slide: {slide}")
                sld = pres.Slides(slide)
                title = ""
                try:
                    if sld.Shapes.HasTitle:
                        title = sld.Shapes.Title.TextFrame.TextRange.Text
                except Exception:
                    pass
                return {"ok": True, "slide": slide, "title": title}
            finally:
                try:
                    pres.Close()
                except Exception:
                    pass

    return _with_retry(_do)


def headed_apply(path: str, ops: dict) -> dict:
    target = Path(path).resolve()
    if not target.is_file():
        raise FileNotFoundError(f"file not found: {target}")

    tool = ops.get("tool")
    if tool not in ("set_text", "add_slide"):
        raise HeadedUnavailable(f"headed mode does not support tool: {tool!r}")

    def _on_exhausted(exc):
        _restore_backup_if_any(path, ops)
        return {
            "ok": False,
            "status": "needs_retry",
            "reason": f"file locked or PowerPoint busy: {exc}",
        }

    def _do():
        with _powerpoint() as app:
            pres = app.Presentations.Open(str(target), WithWindow=True)
            try:
                if tool == "add_slide":
                    pres.Slides.Add(pres.Slides.Count + 1, 12)  # 12 = ppLayoutBlank
                    pres.Save()
                    return {"ok": True, "slide": pres.Slides.Count}
                if tool == "set_text":
                    slide_num = int(ops.get("slide", 1))
                    if slide_num < 1 or slide_num > pres.Slides.Count:
                        raise ValueError(f"unknown slide: {slide_num}")
                    sld = pres.Slides(slide_num)
                    shape_name = ops.get("shape", "title")
                    text = ops.get("text", "")
                    if shape_name.lower() == "title" and sld.Shapes.HasTitle:
                        sld.Shapes.Title.TextFrame.TextRange.Text = text
                    else:
                        found = False
                        for idx in range(1, sld.Shapes.Count + 1):
                            s = sld.Shapes(idx)
                            if s.Name == shape_name:
                                s.TextFrame.TextRange.Text = text
                                found = True
                                break
                        if not found:
                            raise ValueError(f"unknown shape: {shape_name!r}")
                    pres.Save()
                    return {"ok": True, "slide": slide_num, "shape": shape_name, "text": text}
            except Exception:
                _restore_backup_if_any(path, ops)
                raise
            finally:
                try:
                    pres.Close()
                except Exception:
                    pass

    return _with_retry(_do, on_retry_exhausted=_on_exhausted)
