"""Preserve / repoint Excel external-link parts across an openpyxl round-trip.

openpyxl rewrites xl/externalLinks/*.xml when it saves -- it drops the XML
declaration and the mc/x14 namespace attributes Excel authored, which makes
Excel flag the cached values as corrupt and offer to "repair" on open. The
cell cache itself is intact, so the repair is harmless, but it is noise.

`finalize` overwrites every external-link part in the saved workbook with the
pristine Excel-authored bytes from the template, so Excel opens it cleanly.
For the consolidated workbook it also applies the month/filename swap to the
subsidiary-file link targets (replacing the old zip-surgery repoint).
"""

import os
import re
import urllib.parse
import zipfile

import cfs_config as cfg

YM_RE = re.compile(r"20(2[5-9]|[3-9]\d)-\d\d")


def finalize(out_path, template_path, prior_ym=None, target_ym=None,
             new_names=None):
    """Restore external-link parts from the template into the saved workbook.

    If target_ym/new_names are given, also repoint subsidiary-file links to
    the new month (used by the consolidated builder). Returns the list of
    repointed target paths.
    """
    with zipfile.ZipFile(template_path) as ztpl:
        ext_parts = {n: ztpl.read(n) for n in ztpl.namelist()
                     if n.startswith("xl/externalLinks/")}

    repointed = []
    if target_ym and new_names:
        for name, buf in list(ext_parts.items()):
            if name.endswith(".rels"):
                text = buf.decode("utf-8")
                m = re.search(r'Target="([^"]+)"', text)
                if not m:
                    continue
                plain = urllib.parse.unquote(m.group(1))
                base = os.path.basename(plain.replace("\\", "/"))
                ent = next((k for k, e in cfg.ENTITIES.items()
                            if base.startswith(f"{e['file_index']}. {prior_ym} CFS {e['file_label']} ")),
                           None)
                if ent:
                    folder = plain[: len(plain) - len(base)]
                    new_plain = YM_RE.sub(target_ym, folder) + new_names[ent]
                    new_t = urllib.parse.quote(new_plain, safe=":/\\.")
                    text = text.replace(m.group(1), new_t)
                    ext_parts[name] = text.encode("utf-8")
                    repointed.append(new_plain)
            else:
                # swap month in cached sheetName values (e.g. "2026-04 CFS USD")
                text = buf.decode("utf-8")
                if YM_RE.search(text):
                    text = re.sub(r'(<sheetName val="[^"]*?)' + YM_RE.pattern + r'([^"]*"/>)',
                                  lambda mm: YM_RE.sub(target_ym, mm.group(0)), text)
                    ext_parts[name] = text.encode("utf-8")

    tmp = out_path + ".tmp"
    with zipfile.ZipFile(out_path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in ext_parts:
                zout.writestr(item, ext_parts[item.filename])
            else:
                zout.writestr(item, zin.read(item.filename))
        # add any external parts openpyxl dropped entirely
        present = set(zin.namelist())
        for name, buf in ext_parts.items():
            if name not in present:
                zout.writestr(name, buf)
    os.replace(tmp, out_path)
    return repointed
