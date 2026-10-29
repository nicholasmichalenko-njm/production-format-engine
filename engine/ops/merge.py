"""Update document styles from a template (the engine's version of Word's CopyStylesFromTemplate).

Rules, chosen to match Word:
- Every style in the template replaces the document's style of the same name, or is added.
- The document keeps its own styleIds, so no paragraph, run or table reference changes.
- basedOn, next and link are remapped by name to the document's ids.
- A style's list numbering comes with it: the template's list definition is imported with new ids and a new
  nsid (never shared with an existing list), and its level styles are remapped.
- docDefaults come from the template. Styles only in the document stay as they are.
Every replaced, added or changed element is marked for the unit guard.
"""
import copy
import random
import re

from lxml import etree
from ..pkg import Package, W, NS
from ..ordered import q, ensure
from .word_ops import _sig, _lvl_signature, sync_linked

RND = random.Random()


def _norm(n):
    return " ".join((n or "").lower().replace(" - ", " ").split())


def _name(s):
    n = s.find("w:name", NS)
    return n.get(q("val")) if n is not None else None


def op_merge_styles(pkg, template_path):
    src = Package(template_path)
    s_st = src.tree(src.word_part("/styles"))
    d_part = pkg.word_part("/styles")
    pkg.touch(d_part)
    d_st = pkg.tree(d_part)
    d_by_name = {_norm(_name(s)): s for s in d_st.findall("w:style", NS)}
    d_ids = {s.get(q("styleId")) for s in d_st.findall("w:style", NS)}
    t_styles = s_st.findall("w:style", NS)
    # template id -> document id
    idmap = {}
    for ts in t_styles:
        tid = ts.get(q("styleId"))
        ds = d_by_name.get(_norm(_name(ts)))
        if ds is not None:
            idmap[tid] = ds.get(q("styleId"))
        else:
            nid, k = tid, 1
            while nid in d_ids:
                k += 1
                nid = f"{tid}{k}"
            idmap[tid] = nid
            d_ids.add(nid)
    # numbering import
    t_nb = src.tree(src.word_part("/numbering")) if src.word_part("/numbering") else None
    d_nbp = pkg.word_part("/numbering")
    numid_map = {}
    report = {"replaced": 0, "added": 0, "lists_imported": 0}

    def import_num(t_numid):
        if t_numid in numid_map:
            return numid_map[t_numid]
        if t_numid == "0" or t_nb is None or d_nbp is None:
            return t_numid
        pkg.touch(d_nbp)
        d_nb = pkg.tree(d_nbp)
        t_num = next((n for n in t_nb.findall("w:num", NS) if n.get(q("numId")) == t_numid), None)
        if t_num is None:
            return "0"
        t_aid = t_num.find("w:abstractNumId", NS).get(q("val"))
        t_abs = next(a for a in t_nb.findall("w:abstractNum", NS) if a.get(q("abstractNumId")) == t_aid)
        # reuse an identical list already in the document (after remapping level styles), so a rerun adds nothing
        probe = copy.deepcopy(t_abs)
        for ps in probe.iter(q("pStyle")):
            ps.set(q("val"), idmap.get(ps.get(q("val")), ps.get(q("val"))))
        want = _lvl_signature(probe)
        if t_num.find("w:lvlOverride", NS) is None:
            for a in d_nb.findall("w:abstractNum", NS):
                if _lvl_signature(a) == want:
                    n_ = next((n for n in d_nb.findall("w:num", NS) if n.find("w:abstractNumId", NS).get(q("val")) == a.get(q("abstractNumId"))
                               and n.find("w:lvlOverride", NS) is None), None)
                    if n_ is not None:
                        numid_map[t_numid] = n_.get(q("numId"))
                        return numid_map[t_numid]
        aid = max([int(a.get(q("abstractNumId"))) for a in d_nb.findall("w:abstractNum", NS)] + [-1]) + 1
        nid = max([int(n.get(q("numId"))) for n in d_nb.findall("w:num", NS)] + [0]) + 1
        nsids = {a.find("w:nsid", NS).get(q("val")) for a in d_nb.findall("w:abstractNum", NS) if a.find("w:nsid", NS) is not None}
        new_abs = copy.deepcopy(t_abs)
        new_abs.set(q("abstractNumId"), str(aid))
        ns_el = new_abs.find("w:nsid", NS)
        nsid = "%08X" % RND.getrandbits(32)
        while nsid in nsids:
            nsid = "%08X" % RND.getrandbits(32)
        if ns_el is None:
            ns_el = ensure(new_abs, "nsid")
        ns_el.set(q("val"), nsid)
        for ps in new_abs.iter(q("pStyle")):
            ps.set(q("val"), idmap.get(ps.get(q("val")), ps.get(q("val"))))
        for tag in ("styleLink", "numStyleLink"):
            el = new_abs.find(f"w:{tag}", NS)
            if el is not None:
                el.set(q("val"), idmap.get(el.get(q("val")), el.get(q("val"))))
        first_num = d_nb.find("w:num", NS)
        (first_num.addprevious(new_abs) if first_num is not None else d_nb.append(new_abs))
        new_num = copy.deepcopy(t_num)
        new_num.set(q("numId"), str(nid))
        new_num.find("w:abstractNumId", NS).set(q("val"), str(aid))
        d_nb.append(new_num)
        pkg.mark(new_abs); pkg.mark(new_num)
        numid_map[t_numid] = str(nid)
        report["lists_imported"] += 1
        return str(nid)

    for ts in t_styles:
        tid = ts.get(q("styleId"))
        new = copy.deepcopy(ts)
        new.set(q("styleId"), idmap[tid])
        for tag in ("basedOn", "next", "link"):
            el = new.find(f"w:{tag}", NS)
            if el is not None:
                el.set(q("val"), idmap.get(el.get(q("val")), el.get(q("val"))))
        for ni in new.iter(q("numId")):
            ni.set(q("val"), import_num(ni.get(q("val"))))
        old = d_by_name.get(_norm(_name(ts)))
        if old is not None:
            if _sig(old) == _sig(new):
                continue
            pkg.mark(old)
            old.addprevious(new)
            d_st.remove(old)
            report["replaced"] += 1
        else:
            d_st.append(new)
            report["added"] += 1
        pkg.mark(new)
    # linked character styles follow their paragraph styles, as Word keeps them
    for ps in d_st.findall("w:style", NS):
        sync_linked(pkg, d_st, ps)
    # document defaults
    t_dd = s_st.find("w:docDefaults", NS)
    d_dd = d_st.find("w:docDefaults", NS)
    if t_dd is not None and (d_dd is None or _sig(t_dd) != _sig(d_dd)):
        new_dd = copy.deepcopy(t_dd)
        if d_dd is not None:
            pkg.mark(d_dd); d_dd.addprevious(new_dd); d_st.remove(d_dd)
        else:
            d_st.insert(0, new_dd)
        pkg.mark(new_dd)
    report["checks"] = [{"expr": "m5_compare", "expect": "word"}]
    return report
