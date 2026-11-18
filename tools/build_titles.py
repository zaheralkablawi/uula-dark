# Builds one small single-slide package per title layout of the UULA video template
import copy, sys
from pptx import Presentation
SRC = 't.pptx'
OUT = '/home/claude/uula-dark/templates/'
TYPES = {'explain': [1, 2, 3], 'better': [4, 5, 6], 'bank': [13, 14, 15], 'practice': [16, 17, 18]}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
for typ, nums in TYPES.items():
    for k, n in enumerate(nums, 1):
        prs = Presentation(SRC)
        part = f'/ppt/slideLayouts/slideLayout{n}.xml'
        lst = prs.slides._sldIdLst
        for sid in list(lst):
            prs.part.drop_rel(sid.get(RID)); lst.remove(sid)
        keep_m = None
        for m in prs.slide_masters:
            for l in m.slide_layouts:
                if str(l.part.partname) == part: keep_m, keep_l = m, l
        # keep only this master and this layout
        mlst = prs.part._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}sldMasterIdLst')
        for mid in list(mlst):
            if prs.part.related_part(mid.get(RID)) is not keep_m.part:
                prs.part.drop_rel(mid.get(RID)); mlst.remove(mid)
        llst = keep_m._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}sldLayoutIdLst')
        for lid in list(llst):
            if keep_m.part.related_part(lid.get(RID)) is not keep_l.part:
                keep_m.part.drop_rel(lid.get(RID)); llst.remove(lid)
        prs.slides.add_slide(keep_l)
        prs.save(OUT + f'title-{typ}-{k}.pptx')
        print(typ, k, 'ok')
