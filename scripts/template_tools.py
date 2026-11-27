#!/usr/bin/env python3
"""Inspect a real template, bind explicit roles, render references and record review."""
import argparse
import copy
import json
import math
import shutil
import subprocess
import tempfile
from pathlib import Path
from types import SimpleNamespace
from datetime import datetime, timezone
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from deck_pipeline import (ROOT, KINDS, dump, digest, digest_file, replace_text, contact_sheet,
                            clone_slide, remove_originals, native_fingerprint, render)


def inspect(args):
    prs = Presentation(args.pptx)
    out = {'slide_size': [prs.slide_width, prs.slide_height], 'slides': []}
    for i, slide in enumerate(prs.slides, 1):
        shapes = []
        for s in slide.shapes:
            item = {'id': s.shape_id, 'name': s.name, 'type': str(s.shape_type),
                    'box': [round(s.left/prs.slide_width, 6), round(s.top/prs.slide_height, 6),
                            round(s.width/prs.slide_width, 6), round(s.height/prs.slide_height, 6)]}
            if s.has_text_frame:
                item['text'] = s.text
            shapes.append(item)
        out['slides'].append({'slide_index': i, 'layout': slide.slide_layout.name, 'shapes': shapes})
    dump(args.out, out)
    print(json.dumps({'slides': len(prs.slides), 'inspection': str(Path(args.out).resolve())}))


def text_shapes(shapes):
    """Yield text-bearing shapes in reading order, descending into groups."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from text_shapes(shape.shapes)
        elif shape.has_text_frame:
            yield shape


CM_PER_EMU = 360000
PT_PER_CM = 72 / 2.54
H_MARGIN_CM = 0.25
V_MARGIN_CM = 0.08
CAPACITY_SLACK = 1.2
DEFAULT_SIZE_PT = 18.0
A_NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P_NS = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

STYLE_BUCKETS = {'TITLE': 'titleStyle', 'CENTER_TITLE': 'titleStyle',
                 'VERTICAL_TITLE': 'titleStyle',
                 'BODY': 'bodyStyle', 'SUBTITLE': 'bodyStyle', 'OBJECT': 'bodyStyle',
                 'VERTICAL_BODY': 'bodyStyle', 'VERTICAL_OBJECT': 'bodyStyle'}


def paragraph_size_pt(para):
    """Explicit size on any run, else the paragraph's own value."""
    size = next((run.font.size for run in para.runs if run.font.size is not None), None)
    return (size or para.font.size).pt if (size or para.font.size) is not None else None


def lststyle_size_pt(element, level):
    """Size declared in a txBody's <a:lstStyle><a:lvlNpPr><a:defRPr>."""
    body = element if element.tag == P_NS + 'txBody' else element.find('.//' + P_NS + 'txBody')
    if body is None:
        return None
    lst = body.find(A_NS + 'lstStyle')
    if lst is None:
        return None
    lvl = lst.find(f'{A_NS}lvl{level + 1}pPr')
    rpr = lvl.find(A_NS + 'defRPr') if lvl is not None else None
    size = rpr.get('sz') if rpr is not None else None
    return int(size) / 100 if size else None


def txstyle_size_pt(master, placeholder_type, level):
    """Size from the master's <p:txStyles>, bucketed by placeholder type."""
    styles = master._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}txStyles')
    if styles is None:
        return None
    bucket = STYLE_BUCKETS.get(str(placeholder_type).split()[0] if placeholder_type else '')
    if bucket is None:
        return None
    style = styles.find(f'{P_NS}{bucket}')
    if style is None:
        return None
    lvl = style.find(f'{A_NS}lvl{level + 1}pPr')
    rpr = lvl.find(A_NS + 'defRPr') if lvl is not None else None
    size = rpr.get('sz') if rpr is not None else None
    return int(size) / 100 if size else None


def inherited_size_pt(shape, slide, level=0):
    """Walk layout placeholder, master placeholder and master txStyles for a size.

    Ordinary text boxes carry their own size; placeholders often inherit it, and
    reading only the run would under-report the size a reader actually sees.
    """
    if not getattr(shape, 'is_placeholder', False):
        return None
    ph_idx = shape.placeholder_format.idx
    ph_type = shape.placeholder_format.type
    layout, master = slide.slide_layout, slide.slide_layout.slide_master
    for container in (layout, master):
        for ph in container.placeholders:
            if ph.placeholder_format.idx == ph_idx:
                size = lststyle_size_pt(ph._element, level)
                if size:
                    return size
    return txstyle_size_pt(master, ph_type, level)


def effective_size_pt(para, shape=None, slide=None):
    """Resolve the size a reader will see, following the OOXML inheritance chain."""
    size = paragraph_size_pt(para)
    if size is None and shape is not None and slide is not None:
        size = inherited_size_pt(shape, slide, getattr(para, 'level', 0) or 0)
    return round(size * 2) / 2 if size is not None else DEFAULT_SIZE_PT


def slot_capacity(shape, para, slide=None):
    """Estimate how much text fits: chars per line, lines, and a slack budget."""
    size = effective_size_pt(para, shape, slide)
    frame = shape.text_frame
    autofit = frame.auto_size == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    w_cm = shape.width / CM_PER_EMU
    h_cm = shape.height / CM_PER_EMU
    if w_cm < 1.0 or h_cm < 0.3:
        return {'capacity_unknown': True, 'font_size_pt': size, 'autofit': autofit}
    wrap = frame.word_wrap
    wrap = True if wrap is None else bool(wrap)
    cpl = max(1, math.floor(max(0.0, w_cm - H_MARGIN_CM) * PT_PER_CM / size))
    lines = 1 if not wrap else max(1, math.floor(max(0.0, h_cm - V_MARGIN_CM) * PT_PER_CM / size))
    return {'chars_per_line': cpl, 'max_lines': lines,
            'budget': max(1, math.floor(cpl * lines * CAPACITY_SLACK)),
            'font_size_pt': size, 'autofit': autofit}


def index_template(pptx):
    """List every editable text paragraph of a template as an addressable slot."""
    prs = Presentation(str(pptx))
    pages = []
    sizes = {}
    for i, slide in enumerate(prs.slides, 1):
        text_slots = []
        for shape in text_shapes(slide.shapes):
            for pi, para in enumerate(shape.text_frame.paragraphs):
                text = ''.join(run.text for run in para.runs)
                if not text.strip():
                    continue
                capacity = slot_capacity(shape, para, slide)
                sizes[capacity['font_size_pt']] = sizes.get(capacity['font_size_pt'], 0) + 1
                text_slots.append({
                    'id': f's{i}_sh{shape.shape_id}_p{pi}',
                    'current_text': text,
                    'address': {'shape_id': shape.shape_id, 'paragraph': pi},
                    'editable': True,
                    'capacity': capacity,
                })
        pages.append({'slide_index': i, 'text_slots': text_slots})
    ranked = sorted(sizes.items(), key=lambda item: -item[0])
    level_of = {size: rank + 1 for rank, (size, _) in enumerate(ranked)}
    for page in pages:
        for slot in page['text_slots']:
            slot['level'] = level_of[slot['capacity']['font_size_pt']]
    return {
        'schema_version': 1,
        'source_pptx': Path(pptx).name,
        'source_sha256': digest_file(Path(pptx)),
        'slide_size_cm': [round(prs.slide_width / CM_PER_EMU, 2),
                          round(prs.slide_height / CM_PER_EMU, 2)],
        'type_scale': [{'level': rank + 1, 'size_pt': size, 'slot_count': count}
                       for rank, (size, count) in enumerate(ranked)],
        'pages': pages,
    }


def index(args):
    data = index_template(args.pptx)
    dump(args.out, data)
    slots = sum(len(page['text_slots']) for page in data['pages'])
    print(json.dumps({'slots': slots, 'index': str(Path(args.out).resolve())}, ensure_ascii=False))


PARAGRAPH_TEXT_TAGS = (A_NS + 'r', A_NS + 'br', A_NS + 'fld')


def replace_paragraph(para, text):
    """Replace one paragraph's text while keeping its paragraph and first-run formatting.

    Runs, soft line breaks and field references are cleared, so no stale glyph
    survives from the template paragraph. Paragraph properties and the trailing
    endParaRPr are kept.
    """
    rpr = None
    if para.runs and para.runs[0]._r.rPr is not None:
        rpr = copy.deepcopy(para.runs[0]._r.rPr)
    for child in list(para._p):
        if child.tag in PARAGRAPH_TEXT_TAGS:
            para._p.remove(child)
    run = para.add_run()
    run.text = text
    if rpr is not None:
        run._r.insert(0, copy.deepcopy(rpr))


def visual_width(text):
    """Width in em units: CJK 1.0, latin/digit 0.5, space 0.35, other 0.8."""
    width = 0.0
    for char in text:
        if '\u4e00' <= char <= '\u9fff' or '\u3000' <= char <= '\u303f' or '\uff00' <= char <= '\uffef':
            width += 1.0
        elif char == ' ':
            width += 0.35
        elif char.isascii():
            width += 0.5
        else:
            width += 0.8
    return width


def slot_index(slots):
    return {(page['slide_index'], slot['id']): slot
            for page in slots['pages'] for slot in page['text_slots']}


PLACEHOLDER_MARKERS = ('Lorem', 'Key Words', '项目名称', 'Question 1', 'Vivamus', '占位')


def placeholder_in(text):
    return next((marker for marker in PLACEHOLDER_MARKERS if marker in text), None)


def build_deck(plan, slots, template, out):
    """Assemble selected template pages in plan order and replace their slot text."""
    if slots.get('source_sha256') != digest_file(Path(template)):
        raise ValueError('槽位索引的来源文件与当前模板不一致，请对当前模板重新运行 index')
    prs = Presentation(str(template))
    originals = len(prs.slides)
    index_by_slot = slot_index(slots)
    warnings = []
    built = []
    for item in plan['slides']:
        if not 1 <= item['source_slide'] <= originals:
            raise ValueError(f'source_slide {item["source_slide"]} 超出模板页数 1..{originals}')
        source_slide = prs.slides[item['source_slide'] - 1]
        provided = {edit['slot'] for edit in item.get('edits', [])}
        required = {slot_id for (slide_index, slot_id), slot in index_by_slot.items()
                    if slide_index == item['source_slide'] and slot['editable']}
        missing = sorted(required - provided)
        if missing:
            raise ValueError(f'slide {item["source_slide"]}: {len(missing)} 个可编辑槽位未覆盖 '
                             f'{missing[:3]}，留着会显示模板占位文字')
        slide = clone_slide(prs, source_slide)
        shapes = {shape.shape_id: shape for shape in text_shapes(slide.shapes)}
        for edit in item.get('edits', []):
            marker = placeholder_in(edit['text'])
            if marker:
                raise ValueError(f'slide {item["source_slide"]} {edit["slot"]}: '
                                 f'文本仍是占位内容 {marker!r}，需替换为真实内容')
            slot = index_by_slot.get((item['source_slide'], edit['slot']))
            if slot is None:
                raise ValueError(f'slide {item["source_slide"]}: 槽位 {edit["slot"]} '
                                 f'不在索引中，检查 edits 与 slots.json 是否同一版本')
            address = slot['address']
            shape = shapes[address['shape_id']]
            replace_paragraph(shape.text_frame.paragraphs[address['paragraph']], edit['text'])
            if edit['text'].rstrip().endswith(('...', '…', '等等')):
                warnings.append(f'slide {item["source_slide"]} {edit["slot"]}: '
                                f'结尾疑似省略号截断，改成完整表述')
            capacity = slot['capacity']
            budget = capacity.get('budget')
            if budget and not capacity.get('autofit'):
                width = visual_width(edit['text'].replace('\n', ''))
                if width > budget:
                    warnings.append(f'slide {item["source_slide"]} {edit["slot"]}: '
                                    f'文本偏长 {width:.0f} > 容量 {budget}，可用更精炼的措辞')
        built.append(slide)
    remove_originals(prs, originals)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return {'pages': len(built), 'warnings': warnings}


def build(args):
    plan = json.loads(Path(args.plan).read_text(encoding='utf-8'))
    slots_path = Path(args.slots)
    slots = json.loads(slots_path.read_text(encoding='utf-8'))
    template = Path(args.template) if args.template else slots_path.parent / slots['source_pptx']
    report = build_deck(plan, slots, template, args.out)
    print(json.dumps({'pages': report['pages'], 'warnings': len(report['warnings']),
                      'out': str(Path(args.out).resolve())}, ensure_ascii=False))
    for warning in report['warnings']:
        print('WARN ' + warning)
    return 0


def walk_shapes(shapes):
    """Yield every shape on a slide, descending into groups."""
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes)


def top_level_ids_touching(shapes, edited):
    """Top-level shape ids whose subtree contains an edited slot.

    ``native_fingerprint`` compares whole top-level shapes, so a group that holds
    one edited slot cannot be split; the group itself is left out instead.
    """
    touching = set()
    for shape in shapes:
        if any(child.shape_id in edited for child in walk_shapes([shape])):
            touching.add(shape.shape_id)
    return touching


def master_signature(slide):
    """Master part bytes plus its theme and image parts, to catch master drift."""
    master = slide.slide_layout.slide_master.part
    targets = [rel.target_part.blob for rel in master.rels.values()
               if not rel.is_external and rel.reltype.endswith(('/theme', '/image'))]
    return master.blob, targets


def verify_deck(plan, slots, template, pptx):
    """Check a built deck still carries its template's fixed objects and planned text.

    ``plan`` may be a path or an already-loaded dict.
    """
    plan_path = Path(plan) if isinstance(plan, (str, Path)) else None
    plan_data = plan if isinstance(plan, dict) else json.loads(plan_path.read_text(encoding='utf-8'))
    errors = []
    original = Presentation(str(template))
    actual = Presentation(str(pptx))
    if (actual.slide_width, actual.slide_height) != (original.slide_width, original.slide_height):
        errors.append('Slide dimensions differ from template')
    if len(actual.slides) != len(plan_data['slides']):
        errors.append(f'Page count mismatch: built {len(actual.slides)}, planned {len(plan_data["slides"])}')
    index_by_slot = slot_index(slots)
    pages = []
    for position, (item, slide) in enumerate(zip(plan_data['slides'], actual.slides), 1):
        page_errors = []
        source_slide = original.slides[item['source_slide'] - 1]
        edited = {index_by_slot[(item['source_slide'], edit['slot'])]['address']['shape_id']
                  for edit in item.get('edits', [])}
        fixed = {shape.shape_id for shape in source_slide.shapes}
        fixed -= top_level_ids_touching(source_slide.shapes, edited)
        if native_fingerprint(source_slide, included=fixed) != native_fingerprint(slide, included=fixed):
            page_errors.append('fixed template objects changed')
        if master_signature(source_slide) != master_signature(slide):
            page_errors.append('slide layout, master or theme changed')
        shapes = {shape.shape_id: shape for shape in text_shapes(slide.shapes)}
        for edit in item.get('edits', []):
            address = index_by_slot[(item['source_slide'], edit['slot'])]['address']
            shape = shapes.get(address['shape_id'])
            if shape is None:
                page_errors.append(f'slot shape missing: {edit["slot"]}')
                continue
            paragraph = shape.text_frame.paragraphs[address['paragraph']]
            if ''.join(run.text for run in paragraph.runs) != edit['text']:
                page_errors.append(f'slot text mismatch: {edit["slot"]}')
        for shape in text_shapes(slide.shapes):
            marker = placeholder_in(shape.text_frame.text)
            if marker:
                page_errors.append(f'成品仍含占位文字 {marker!r}: shape {shape.shape_id}')
        errors.extend(f'第{position}页: {error}' for error in page_errors)
        pages.append({'page': position, 'source_slide': item['source_slide'],
                      'passed': not page_errors, 'errors': page_errors})
    return {
        'mode': 'page-slot',
        'pptx_sha256': digest_file(Path(pptx)),
        'plan_sha256': digest(plan_data),
        'template': {'source_sha256': slots.get('source_sha256')},
        'structure': {'passed': not errors, 'errors': errors, 'pages': pages},
        'visual_review': {'status': 'pending'},
    }


def accept_deck(plan, slots, template, pptx, output, *, renders=None, reviewer=None, comment=None):
    """Record structure checks and rendered-page evidence for a slot-built deck."""
    if (reviewer is None) != (comment is None) or (reviewer is not None and not (reviewer.strip() and comment.strip())):
        raise ValueError('Provide a non-empty reviewer and comment together')
    if reviewer and renders is None:
        raise ValueError('Review existing final renders with --renders before recording visual acceptance')
    output = Path(output)
    qa, sheet = output / 'qa.json', output / 'contact-sheet.jpg'
    if {Path(plan).resolve(), Path(pptx).resolve()} & {qa.resolve(), sheet.resolve()}:
        raise ValueError('Acceptance output must be separate from its inputs')
    report = verify_deck(plan, slots, template, pptx)
    dump(qa, report)
    if not report['structure']['passed']:
        raise ValueError('\n'.join(report['structure']['errors']))
    if renders is None:
        renders = Path(tempfile.mkdtemp(prefix='renders-', dir=output))
        render(SimpleNamespace(pptx=pptx, out=renders))
    renders = Path(renders)
    proof = json.loads((renders / 'render.json').read_text(encoding='utf-8'))
    if proof['pptx_sha256'] != report['pptx_sha256']:
        raise ValueError('Render belongs to another revision of the deck; re-render the final file')
    images = {p.name: digest_file(p) for p in renders.glob('*.png')}
    renderer = proof.get('renderer')
    if (not isinstance(renderer, str) or not renderer.strip() or not images
        or proof['pages'] != len(images) or len(images) != len(report['structure']['pages'])
        or images != proof['image_sha256']):
        raise ValueError('Render provenance, page count or image hashes do not match')
    contact_sheet(SimpleNamespace(images=str(renders), out=str(sheet), columns=3))
    report['render_evidence'] = {**proof, 'directory': str(renders.resolve()),
                                 'contact_sheet_sha256': digest_file(sheet)}
    if reviewer:
        report['visual_review'] = {'status': 'passed', 'reviewer': reviewer, 'comment': comment}
    dump(qa, report)
    return report


def accept(args):
    slots_path = Path(args.slots)
    slots = json.loads(slots_path.read_text(encoding='utf-8'))
    template = Path(args.template) if args.template else slots_path.parent / slots['source_pptx']
    report = accept_deck(args.plan, slots, template, args.pptx, args.out,
                         renders=Path(args.renders) if args.renders else None,
                         reviewer=args.reviewer, comment=args.comment)
    print(json.dumps({'structure_passed': report['structure']['passed'],
                      'visual_review': report['visual_review']['status'],
                      'pages': len(report['structure']['pages']),
                      'qa': str((Path(args.out) / 'qa.json').resolve())}, ensure_ascii=False))


def check_config(config, prs):
    if abs(prs.slide_width/prs.slide_height/(16/9)-1) > 0.002:
        raise ValueError('Template is not 16:9. Do not silently alter the original dimensions.')
    masters = config.get('masters', {})
    if not {'cover', 'body'} <= set(masters):
        raise ValueError('At least cover and body masters must be bound')
    allowed_roles = {'title', 'subtitle', 'page_number', 'footer', 'closing_text', 'navigation', 'navigation_1', 'navigation_2', 'navigation_3'}
    for kind, m in masters.items():
        if kind not in KINDS:
            raise ValueError('Unknown master: ' + kind)
        idx = m.get('slide_index')
        if not isinstance(idx, int) or not 1 <= idx <= len(prs.slides):
            raise ValueError(kind + ': slide_index must be 1-based within template')
        slide = prs.slides[idx - 1]
        ids = {s.shape_id: s for s in slide.shapes}
        roles, removal = m.get('role_shapes', {}), m.get('remove_shape_ids', [])
        if not set(roles) <= allowed_roles:
            raise ValueError(kind + ': unsupported role')
        if not isinstance(removal, list) or any(not isinstance(x, int) for x in removal):
            raise ValueError(kind + ': remove_shape_ids must be integer list')
        if len(set(roles.values())) != len(roles):
            raise ValueError(kind + ': one text shape cannot serve multiple roles')
        if set(roles.values()) & set(removal):
            raise ValueError(kind + ': dynamic text role also selected for deletion')
        if kind == 'body':
            required = {'title', 'subtitle'}
            if m.get('page_number_mode') == 'inherited-field':
                field_id = m.get('layout_page_number_shape_id')
                inherited = next((s for s in slide.slide_layout.shapes if s.shape_id == field_id), None)
                if inherited is None or not inherited._element.xpath('.//a:fld[@type="slidenum"]'):
                    raise ValueError('Inherited page number must identify the real layout slidenum field')
            else:
                required.add('page_number')
            if not required <= set(roles):
                raise ValueError('body requires title, subtitle and a native page-number role or inherited-field')
        if kind == 'cover' and 'title' not in roles:
            raise ValueError('cover requires title role')
        for role, sid in roles.items():
            if sid not in ids or not ids[sid].has_text_frame:
                raise ValueError(kind + ': role must identify a top-level original text shape: ' + role)
        if not set(removal) <= set(ids):
            raise ValueError(kind + ': unknown removal shape IDs')
        if kind == 'body':
            box = m.get('content_box')
            if not isinstance(box, list) or len(box) != 4 or any(not isinstance(x, (int,float)) for x in box):
                raise ValueError('body content_box must be [x,y,w,h] normalized numbers')
            x,y,w,h = box
            if min(x,y) < 0 or min(w,h) <= 0 or x+w > 1 or y+h > 1:
                raise ValueError('body content box outside slide')
            for zone in m.get('transparent_boxes', []):
                if len(zone) != 4 or min(zone[:2]) < 0 or min(zone[2:]) <= 0 or zone[0]+zone[2] > 1 or zone[1]+zone[3] > 1:
                    raise ValueError('Transparent watermark box outside slide')
            # Dynamic header/page tab must not be obscured by the opaque content image.
            for role, sid in roles.items():
                s = ids[sid]
                sx,sy,sw,sh = s.left/prs.slide_width,s.top/prs.slide_height,s.width/prs.slide_width,s.height/prs.slide_height
                if max(x,sx) < min(x+w,sx+sw) and max(y,sy) < min(y+h,sy+sh):
                    raise ValueError('body content box overlaps text role: ' + role)
    return masters


def bind(args):
    src, cfg = Path(args.pptx).resolve(), Path(args.config).resolve()
    if not src.is_file():
        raise ValueError('Original PPTX not found')
    config = json.loads(cfg.read_text(encoding='utf-8'))
    prs = Presentation(str(src))
    masters = check_config(config, prs)
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    destination = out / 'source.pptx'
    if destination.exists() and digest_file(destination) != digest_file(src) and not args.replace:
        raise ValueError('A different bound source exists. Preserve it or explicitly use --replace after user requests a template change.')
    if src != destination:
        shutil.copy2(src, destination)
    # Raw and blank previews are rendered separately; the original source stays untouched.
    blank = Presentation(str(destination))
    selected = {m['slide_index']: m for m in masters.values()}
    for idx, m in selected.items():
        slide = blank.slides[idx-1]
        for s in list(slide.shapes):
            if s.shape_id in m.get('remove_shape_ids', []):
                s._element.getparent().remove(s._element)
            elif s.shape_id in m.get('role_shapes', {}).values():
                replace_text(s, '')
        if slide.has_notes_slide:
            slide.notes_slide.notes_text_frame.text = ''
    blank.save(str(out / 'blank-preview.pptx'))
    for kind, m in masters.items():
        source_slide = prs.slides[m['slide_index']-1]
        m['role_boxes'] = {}
        for role, sid in m.get('role_shapes', {}).items():
            sh = next(x for x in source_slide.shapes if x.shape_id == sid)
            m['role_boxes'][role] = [sh.left/prs.slide_width, sh.top/prs.slide_height,
                                     sh.width/prs.slide_width, sh.height/prs.slide_height]
        m['reference_image'] = f'references/{kind}.png'
        m['blank_reference_image'] = f'blank/{kind}.png'
    profile = {'schema_version': 1, 'source_pptx': 'source.pptx',
               'source_sha256': digest_file(destination), 'masters': masters,
               'visual_status': 'pending', 'provenance': config.get('provenance', 'User supplied original PPTX'),
               'bound_at': datetime.now(timezone.utc).isoformat()}
    dump(out / 'profile.json', profile)
    if out == (ROOT / 'assets/master').resolve():
        dump(ROOT / 'assets/template-status.json', {'status': 'bound-awaiting-visual-review',
                                                  'profile': 'master/profile.json'})
    print(json.dumps({'profile': str(out / 'profile.json'), 'visual_status': 'pending'}, ensure_ascii=False))


def executable(name):
    found = shutil.which(name)
    if not found:
        raise ValueError(f'{name} not on PATH. Use the host presentation renderer or add an installed executable to PATH.')
    return found


def render(args):
    profile_path = Path(args.profile).resolve()
    p = json.loads(profile_path.read_text(encoding='utf-8'))
    base = profile_path.parent
    src = base / p['source_pptx']
    if digest_file(src) != p['source_sha256']:
        raise ValueError('Source template changed')
    lo, pp = executable('soffice'), executable('pdftoppm')
    with tempfile.TemporaryDirectory(prefix='ppt-master-render-') as temp:
        t = Path(temp)
        for source_name, output_dir in ((p['source_pptx'], 'references'), ('blank-preview.pptx', 'blank')):
            dest = t / output_dir
            dest.mkdir()
            r = subprocess.run([lo, f'-env:UserInstallation={ (t / (output_dir+"-lo-profile")).as_uri() }',
                                '--headless','--convert-to','pdf','--outdir',str(dest),str(base/source_name)],
                               capture_output=True, text=True, timeout=180)
            pdf = dest / (Path(source_name).stem + '.pdf')
            if r.returncode or not pdf.is_file():
                raise ValueError('Template render failed: ' + (r.stderr + r.stdout)[-1500:])
            (base / output_dir).mkdir(exist_ok=True)
            for kind, m in p['masters'].items():
                idx = m['slide_index']
                prefix = dest / kind
                subprocess.run([pp,'-f',str(idx),'-l',str(idx),'-singlefile','-scale-to-x','1920','-scale-to-y','1080',
                                '-png',str(pdf),str(prefix)],check=True,capture_output=True,timeout=90)
                shutil.copy2(prefix.with_suffix('.png'), base / output_dir / (kind + '.png'))
    print(json.dumps({'references': [str(base / m['reference_image']) for m in p['masters'].values()],
                      'review': 'Compare originals and blank masters visually before mark-verified'}))
    warnings = []
    if shutil.which('fc-list'):
        fonts = subprocess.run(['fc-list', ':lang=zh'], capture_output=True, text=True, timeout=20)
        if fonts.returncode == 0 and not fonts.stdout.strip():
            warnings.append('Chinese fonts missing: reference text renders blank; fixed geometry only was previewed. Re-render with original licensed fonts or a suitable host renderer before acceptance.')
    p['visual_status'] = 'pending'
    p.pop('visual_review', None)
    p['render_warnings'] = warnings
    p['reference_renderer'] = 'LibreOffice + Poppler'
    dump(profile_path, p)
    if profile_path == ROOT / 'assets/master/profile.json':
        dump(ROOT / 'assets/template-status.json',
             {'status': 'bound-awaiting-visual-review', 'profile': 'master/profile.json'})


def mark_verified(args):
    path = Path(args.profile).resolve()
    p = json.loads(path.read_text(encoding='utf-8'))
    src = path.parent / p['source_pptx']
    if digest_file(src) != p['source_sha256']:
        raise ValueError('Source changed')
    if p.get('render_warnings'):
        raise ValueError('Template text render unresolved: ' + '; '.join(p['render_warnings']))
    hashes = {}
    for kind,m in p['masters'].items():
        for key in ('reference_image','blank_reference_image'):
            fp = path.parent / m[key]
            if not fp.is_file():
                raise ValueError('Render references first; missing ' + str(fp))
            hashes[kind + ':' + key] = digest_file(fp)
    if not args.comment.strip():
        raise ValueError('Concrete visual review comment required')
    p.update(visual_status='verified', visual_review={'reviewer':args.reviewer,'comment':args.comment,
                                                   'reference_sha256':hashes,
                                                   'reviewed_at':datetime.now(timezone.utc).isoformat()})
    dump(path,p)
    if path == ROOT / 'assets/master/profile.json':
        dump(ROOT / 'assets/template-status.json', {'status':'ready','profile':'master/profile.json'})
    print('Template visual review recorded; persist assets with skill commit/push.')


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    sp = ap.add_subparsers(dest='command',required=True)
    p=sp.add_parser('inspect');p.add_argument('--pptx',required=True);p.add_argument('--out',required=True);p.set_defaults(func=inspect)
    p=sp.add_parser('index');p.add_argument('--pptx',required=True);p.add_argument('--out',required=True);p.set_defaults(func=index)
    p=sp.add_parser('build');p.add_argument('--plan',required=True);p.add_argument('--slots',required=True);p.add_argument('--out',required=True);p.add_argument('--template',default=None);p.set_defaults(func=build)
    p=sp.add_parser('accept');p.add_argument('--plan',required=True);p.add_argument('--slots',required=True);p.add_argument('--pptx',required=True);p.add_argument('--out',required=True);p.add_argument('--template',default=None);p.add_argument('--renders',default=None);p.add_argument('--reviewer',default=None);p.add_argument('--comment',default=None);p.set_defaults(func=accept)
    p=sp.add_parser('bind');p.add_argument('--pptx',required=True);p.add_argument('--config',required=True);p.add_argument('--out',default=str(ROOT/'assets/master'));p.add_argument('--replace',action='store_true');p.set_defaults(func=bind)
    p=sp.add_parser('render');p.add_argument('--profile',required=True);p.set_defaults(func=render)
    p=sp.add_parser('mark-verified');p.add_argument('--profile',required=True);p.add_argument('--reviewer',required=True);p.add_argument('--comment',required=True);p.set_defaults(func=mark_verified)
    args=ap.parse_args()
    try:args.func(args)
    except (ValueError,OSError,KeyError,subprocess.SubprocessError) as e:ap.exit(2,str(e)+'\n')


if __name__=='__main__':main()
