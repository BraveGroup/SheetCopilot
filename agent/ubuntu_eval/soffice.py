"""
Headless LibreOffice (soffice) instance manager and UNO bootstrap.

The LibreOffice-UNO engine needs a running ``soffice`` process exposing a UNO
bridge.  For multiprocessing each worker process owns its *own* private
``soffice`` instance -- a separate user profile and a unique named pipe -- so the
workers never contend on a shared office instance (a classic source of UNO
flakiness).

Typical use (one instance per process, reused across many comparisons)::

    from ubuntu_eval.soffice import get_process_instance
    inst = get_process_instance()          # lazily started, cached per-process
    doc = inst.load("/abs/path/file.xlsx")  # hidden, read-only, recalculated
    ...
    doc.close(False)

``import uno`` requires the ``python3-uno`` package (installed by
``setup_ubuntu_eval.sh``).  We import it lazily so that the openpyxl-only path
keeps working on machines without LibreOffice.
"""

import os
import subprocess
import time

_DEVNULL = open(os.devnull, "wb")


def _find_soffice():
    for cand in ("soffice", "libreoffice"):
        from shutil import which

        p = which(cand)
        if p:
            return p
    for cand in (
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
    ):
        if os.path.exists(cand):
            return cand
    raise RuntimeError(
        "Could not find 'soffice'. Install LibreOffice (see setup_ubuntu_eval.sh)."
    )


class SofficeInstance:
    """A single headless soffice process plus its UNO desktop connection."""

    def __init__(self, pipe_name=None, profile_dir=None, start_timeout=60.0):
        pid = os.getpid()
        self.pipe_name = pipe_name or f"scuno_{pid}_{id(self)}"
        self.profile_dir = profile_dir or f"/tmp/sc_lo_profile_{pid}_{id(self)}"
        self.start_timeout = start_timeout
        self.proc = None
        self.ctx = None
        self.smgr = None
        self.desktop = None

    # -- lifecycle -------------------------------------------------------- #
    def start(self):
        if self.desktop is not None:
            return self
        soffice = _find_soffice()
        os.makedirs(self.profile_dir, exist_ok=True)
        profile_url = "file://" + os.path.abspath(self.profile_dir)
        cmd = [
            soffice,
            "--headless",
            "--invisible",
            "--nodefault",
            "--norestore",
            "--nologo",
            "--nofirststartwizard",
            "--nocrashreport",
            f"--accept=pipe,name={self.pipe_name};urp;StarOffice.ComponentContext",
            f"-env:UserInstallation={profile_url}",
        ]
        self.proc = subprocess.Popen(cmd, stdout=_DEVNULL, stderr=_DEVNULL)
        self._connect()
        return self

    def _connect(self):
        import uno  # lazy

        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        url = (
            f"uno:pipe,name={self.pipe_name};urp;StarOffice.ComponentContext"
        )
        deadline = time.time() + self.start_timeout
        last_err = None
        while time.time() < deadline:
            try:
                self.ctx = resolver.resolve(url)
                break
            except Exception as e:  # NoConnectException until the pipe is ready
                last_err = e
                time.sleep(0.5)
        else:
            self.stop()
            raise RuntimeError(f"Timed out connecting to soffice: {last_err}")

        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

    def load(self, path):
        """Load a spreadsheet hidden + read-only and return the document.

        Opening through LibreOffice recalculates formulas and materialises chart
        series / pivot (DataPilot) tables, which is exactly why this engine is
        used for charts and pivot tables."""
        import uno  # lazy
        from com.sun.star.beans import PropertyValue

        if self.desktop is None:
            self.start()
        file_url = uno.systemPathToFileUrl(os.path.abspath(path))
        props = []
        for name, value in (("Hidden", True), ("ReadOnly", True), ("UpdateDocMode", 3)):
            pv = PropertyValue()
            pv.Name = name
            pv.Value = value
            props.append(pv)
        doc = self.desktop.loadComponentFromURL(file_url, "_blank", 0, tuple(props))
        if doc is None:
            raise RuntimeError(f"LibreOffice failed to load {path}")
        try:
            doc.calculateAll()
        except Exception:
            pass
        return doc

    def stop(self):
        try:
            if self.desktop is not None:
                self.desktop.terminate()
        except Exception:
            pass
        self.desktop = None
        self.ctx = None
        self.smgr = None
        if self.proc is not None:
            try:
                self.proc.terminate()
                self.proc.wait(timeout=10)
            except Exception:
                try:
                    self.proc.kill()
                except Exception:
                    pass
            self.proc = None
        # Best-effort profile cleanup.
        try:
            import shutil

            shutil.rmtree(self.profile_dir, ignore_errors=True)
        except Exception:
            pass

    def __enter__(self):
        return self.start()

    def __exit__(self, *exc):
        self.stop()


def convert_to_xlsx(path):
    """Round-trip ``path`` through headless LibreOffice and return the path of the
    re-saved ``.xlsx``.

    Used to repair workbooks whose XML openpyxl cannot parse.  Runs as a one-shot
    CLI conversion with its own throwaway user profile (so it never clashes with
    a worker's long-lived UNO instance)."""
    soffice = _find_soffice()
    pid = os.getpid()
    outdir = f"/tmp/sc_lo_convert_{pid}_{abs(hash(path)) % 100000}"
    os.makedirs(outdir, exist_ok=True)
    profile_url = "file://" + os.path.abspath(outdir + "_profile")
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nologo",
        f"-env:UserInstallation={profile_url}",
        "--convert-to",
        "xlsx",
        "--outdir",
        outdir,
        os.path.abspath(path),
    ]
    subprocess.run(cmd, stdout=_DEVNULL, stderr=_DEVNULL, timeout=180, check=False)
    base = os.path.splitext(os.path.basename(path))[0] + ".xlsx"
    out = os.path.join(outdir, base)
    if not os.path.exists(out):
        raise RuntimeError(f"LibreOffice could not repair {path}")
    return out


# --------------------------------------------------------------------------- #
# Per-process singleton (so a worker reuses one soffice across many tasks)
# --------------------------------------------------------------------------- #
_PROCESS_INSTANCE = None


def get_process_instance():
    global _PROCESS_INSTANCE
    if _PROCESS_INSTANCE is None:
        _PROCESS_INSTANCE = SofficeInstance().start()
    return _PROCESS_INSTANCE


def shutdown_process_instance():
    global _PROCESS_INSTANCE
    if _PROCESS_INSTANCE is not None:
        _PROCESS_INSTANCE.stop()
        _PROCESS_INSTANCE = None
