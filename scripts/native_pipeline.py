#!/usr/bin/env python3
"""Build editable decks and verify built or host-authored PPTX files."""
import argparse
from collections import Counter
import json
import math
from pathlib import Path
import subprocess
import tempfile
from types import SimpleNamespace
from zipfile import BadZipFile, ZipFile

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

import deck_pipeline as shared
from template_tools import check_config

OBJECT_TYPES = {'text', 'shape', 'connector', 'table', 'chart', 'image'}
CHART_TYPES = {'COLUMN_CLUSTERED', 'BAR_CLUSTERED', 'LINE', 'LINE_MARKERS'}


def load_plan(path):
    plan = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(plan, dict) or not isinstance(plan.get('title'), str) or not plan['title'].strip():
        raise ValueError('Deck title is required')
    for field in ('font', 'font_substitution', 'template_profile'):
        if field in plan and not isinstance(plan[field], str):
            raise ValueError(field + ' must be a string')
    pages = plan.get('slides')
    if not isinstance(pages, list) or not pages:
        raise ValueError('slides must be a non-empty list')
    for order, page in enumerate(pages, 1):
        sid = f'P{order:03d}'
        if not isinstance(page, dict) or page.get('id') != sid or page.get('order', order) != order:
            raise ValueError(f'Page {order}: expected ID {sid} and matching order')
        page['order'] = order
        if page.get('kind') not in shared.KINDS:
            raise ValueError(f'{sid}: unsupported kind')
        for key in ('title', 'speaker_notes'):
            if not isinstance(page.get(key), str) or not page[key].strip():
                raise ValueError(f'{sid}: {key} is required')
        for key in ('subtitle', 'footer'):
            if not isinstance(page.get(key, ''), str):
                raise ValueError(f'{sid}: {key} must be a string')
        for key in ('on_slide', 'sources'):
            items = page.setdefault(key, [])
            if not isinstance(items, list) or any(not isinstance(item, str) for item in items):
                raise ValueError(f'{sid}: {key} must be a list of strings')
        if page['kind'] == 'body' and not any(s.strip() for s in page['on_slide']):
            raise ValueError(f'{sid}: body needs planned on_slide text')
        objects = page.get('objects', [])
        if not isinstance(objects, list) or any(
            not isinstance(obj, dict) or obj.get('type') not in OBJECT_TYPES for obj in objects
        ):
            raise ValueError(f'{sid}: invalid objects list')
        required = page.get('required_objects', {})
        if not isinstance(required, dict) or any(
            key not in OBJECT_TYPES or type(value) is not int or value < 0
            for key, value in required.items()
        ):
            raise ValueError(f'{sid}: invalid required_objects')
    profile_path, profile = shared.template(plan, path, required=True)
    source = (profile_path.parent / profile['source_pptx']).resolve()
    return plan, profile_path, profile, source


def apply_font(font, style):
    if name := style.get('font'):
        font.name = name
        east_asian = font._rPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
        if east_asian is None:
            east_asian = OxmlElement('a:ea')
            font._rPr.insert_element_before(east_asian, 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst')
        east_asian.set('typeface', name)
    if 'size' in style:
        font.size = Pt(style['size'])
    if 'bold' in style:
        font.bold = style['bold']
    if 'color' in style:
        font.color.rgb = RGBColor.from_string(style['color'])


def write_text(frame, obj, default_font):
    frame.clear()
    frame.word_wrap = obj.get('wrap', True)
    margins = obj.get('margins', [0, 0, 0, 0])
    frame.margin_left, frame.margin_top, frame.margin_right, frame.margin_bottom = map(Inches, margins)
    if 'anchor' in obj:
        frame.vertical_anchor = MSO_ANCHOR[obj['anchor']]
    style = {'font': default_font, **obj.get('style', {})}
    overrides = obj.get('paragraph_styles', [])
    text = obj.get('text', '')
    for index, line in enumerate(text.split('\n')):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        if 'align' in obj:
            paragraph.alignment = PP_ALIGN[obj['align']]
        if 'space_after' in obj:
            paragraph.space_after = Pt(obj['space_after'])
        run = paragraph.add_run()
        run.text = line
        apply_font(run.font, {**style, **overrides[index]} if index < len(overrides) else style)


def set_line(shape, stroke):
    if 'color' not in stroke:
        shape.line.fill.background()
        return
    shape.line.color.rgb = RGBColor.from_string(stroke['color'])
    shape.line.width = Pt(stroke.get('width', 1))
    if 'dash' in stroke:
        shape.line.dash_style = MSO_LINE_DASH_STYLE[stroke['dash']]


def set_fill(shape, color):
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(color)


def coordinates(obj, key, prs, master):
    values = obj[key]
    if not isinstance(values, list) or len(values) != 4 or any(
        type(v) not in (int, float) or not math.isfinite(v) for v in values
    ):
        raise ValueError(f'{key} must contain four finite numbers in inches')
    x, y, a, b = map(Inches, values)
    if key == 'box' and (a <= 0 or b <= 0):
        raise ValueError('box width and height must be positive')
    x2, y2 = (x + a, y + b) if key == 'box' else (a, b)
    sizes = (prs.slide_width, prs.slide_height)
    bx, by, bw, bh = master.get('content_box', [0, 0, 1, 1])
    if (min(x, x2) < bx * sizes[0] - 2 or max(x, x2) > (bx + bw) * sizes[0] + 2
        or min(y, y2) < by * sizes[1] - 2 or max(y, y2) > (by + bh) * sizes[1] + 2):
        raise ValueError('object is outside the template content box')
    return x, y, a, b


def add_object(slide, obj, prs, master, path, default_font):
    shapes = slide.shapes
    kind = obj['type']
    if kind == 'connector':
        shape = shapes.add_connector(MSO_CONNECTOR[obj.get('line_type', 'STRAIGHT')],
                                     *coordinates(obj, 'points', prs, master))
        set_line(shape, obj.get('stroke', {'color': '172C89'}))
        if obj.get('arrow'):
            arrow = OxmlElement('a:tailEnd')
            arrow.set('type', 'triangle')
            shape._element.spPr.get_or_add_ln().insert_element_before(arrow, 'a:extLst')
        return
    box = coordinates(obj, 'box', prs, master)
    if kind == 'image':
        pic = shapes.add_picture(str(shared.resolve(path, obj['path'])), *box)
        if alt := obj.get('alt'):
            pic._element.nvPicPr.cNvPr.set('descr', alt)
        return
    if kind == 'chart':
        chart_type = obj.get('chart_type', 'COLUMN_CLUSTERED')
        if chart_type not in CHART_TYPES:
            raise ValueError('unsupported native chart type: ' + str(chart_type))
        categories = obj['categories']
        if not categories or any(not isinstance(value, str) or not value.strip() for value in categories):
            raise ValueError('chart requires non-empty category labels')
        data = CategoryChartData()
        data.categories = categories
        if not obj['series']:
            raise ValueError('chart requires at least one series')
        for series in obj['series']:
            if len(series['values']) != len(categories) or any(
                type(value) not in (int, float) or not math.isfinite(value) for value in series['values']
            ):
                raise ValueError('chart series requires one finite value per category')
            data.add_series(series['name'], series['values'])
        chart = shapes.add_chart(XL_CHART_TYPE[chart_type], *box, data).chart
        apply_font(chart.font, {'font': default_font, **obj.get('style', {})})
        chart.has_legend = obj.get('legend', len(obj['series']) > 1)
        if title := obj.get('title'):
            write_text(chart.chart_title.text_frame, {**obj, 'text': title}, default_font)
        return
    if kind == 'table':
        rows = obj['rows']
        if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
            raise ValueError('table requires non-empty rectangular rows')
        table = shapes.add_table(len(rows), len(rows[0]), *box).table
        if widths := obj.get('column_widths'):
            if len(widths) != len(rows[0]) or any(width <= 0 for width in widths) or not math.isclose(sum(widths), 1):
                raise ValueError('column_widths must be positive fractions summing to one')
            for column, fraction in zip(table.columns, widths):
                column.width = round(box[2] * fraction)
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                cell = table.cell(i, j)
                header = obj.get('header', {}) if i == 0 else {}
                write_text(cell.text_frame, {
                    **obj, 'text': str(value), 'style': {**obj.get('style', {}), **header.get('style', {})}
                }, default_font)
                if 'fill' in header or 'fill' in obj:
                    set_fill(cell, header.get('fill', obj.get('fill')))
        return
    if kind == 'shape':
        shape = shapes.add_shape(MSO_SHAPE[obj['shape']], *box)
        set_fill(shape, obj.get('fill'))
        set_line(shape, obj.get('stroke', {}))
    else:
        shape = shapes.add_textbox(*box)
    write_text(shape.text_frame, obj, default_font)


def build(path, output):
    plan, profile_path, profile, source = load_plan(path)
    if profile.get('visual_status') != 'verified' or profile.get('render_warnings'):
        raise ValueError('Template binding must be visually verified before native build')
    hashes = profile.get('visual_review', {}).get('reference_sha256', {})
    for kind, master in profile['masters'].items():
        for key in ('reference_image', 'blank_reference_image'):
            if stamp := hashes.get(f'{kind}:{key}'):
                image = profile_path.parent / master[key]
                if not image.is_file() or shared.digest_file(image) != stamp:
                    raise ValueError(f'Template reference changed since review: {kind}/{key}')
    if output.resolve() in {source, path.resolve(), profile_path}:
        raise ValueError('Output must be separate from the source template, profile and plan')
    prs = Presentation(source)
    masters = check_config(profile, prs)
    originals = list(prs.slides)
    for page in plan['slides']:
        sid = page['id']
        if page['kind'] not in masters:
            raise ValueError(f'{sid}: no bound template page for {page["kind"]}')
        master = masters[page['kind']]
        check_navigation(page, master)
        slide = shared.add_template_page(prs, originals[master['slide_index'] - 1], master, page)
        if font := plan.get('font'):
            for shape in slide.shapes:
                if shape.shape_id in master.get('role_shapes', {}).values():
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            apply_font(run.font, {'font': font})
        for index, obj in enumerate(page.get('objects', []), 1):
            try:
                add_object(slide, obj, prs, master, path, plan.get('font'))
            except (KeyError, TypeError, ValueError, OSError) as error:
                raise ValueError(f'{sid}: object {index} ({obj["type"]}): {error}') from error
        slide.notes_slide.notes_text_frame.text = shared.notes_text(page)
    shared.remove_originals(prs, len(originals))
    prs.core_properties.title = plan['title']
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=output.parent) as temp:
        staged = Path(temp) / 'deck.pptx'
        prs.save(staged)
        report = verify(path, staged)
        if not report['structure']['passed']:
            raise ValueError('\n'.join(report['structure']['errors']))
        staged.replace(output)


def check_navigation(page, master):
    if master.get('navigation_count') and len(page['on_slide']) != master['navigation_count']:
        raise ValueError(f'{page["id"]}: on_slide must match the bound navigation row count')


def object_kind(shape):
    if shape.has_table:
        return 'table'
    if shape.has_chart:
        return 'chart'
    return {
        MSO_SHAPE_TYPE.LINE: 'connector',
        MSO_SHAPE_TYPE.PICTURE: 'image',
        MSO_SHAPE_TYPE.TEXT_BOX: 'text',
        MSO_SHAPE_TYPE.AUTO_SHAPE: 'shape',
    }.get(shape.shape_type)


def visible_text(shape):
    if shape.has_text_frame:
        yield shape.text
    elif shape.has_table:
        yield from (cell.text for cell in shape.table.iter_cells())
    elif shape.has_chart:
        chart = shape.chart
        if chart.has_title:
            yield chart.chart_title.text_frame.text
        for plot in chart.plots:
            yield from (category.label for category in plot.categories)
        if chart.has_legend:
            yield from (series.name for series in chart.series)
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            yield from visible_text(child)


def missing_text(required, actual):
    def normalize(value):
        return ' '.join(value.split())
    present = {normalize(line) for value in actual for line in [value, *value.splitlines()]}
    for value in required:
        text = normalize(value)
        if text and text not in present:
            yield value


def chart_data(chart):
    return (
        chart.chart_type.name,
        [category.label for category in chart.plots[0].categories],
        [(series.name, tuple(series.values)) for series in chart.series],
    )


def master_signature(slide):
    master = slide.slide_layout.slide_master.part
    targets = [rel.target_part.blob for rel in master.rels.values()
               if not rel.is_external and rel.reltype.endswith(('/theme', '/image'))]
    return master.blob, targets


def verify(path, pptx):
    plan, profile_path, profile, source = load_plan(path)
    original = Presentation(source)
    masters = check_config(profile, original)
    with ZipFile(pptx) as archive:
        bad_member = archive.testzip()
    actual = Presentation(pptx)
    errors = []
    if bad_member:
        errors.append('PPTX ZIP integrity failure: ' + bad_member)
    if len(actual.slides) != len(plan['slides']):
        errors.append('Page count mismatch')
    if (actual.slide_width, actual.slide_height) != (original.slide_width, original.slide_height):
        errors.append('Slide dimensions differ from template')
    pages = []
    for page, slide in zip(plan['slides'], actual.slides):
        sid = page['id']
        page_errors = []
        master = masters.get(page['kind'])
        if master is None:
            errors.append(f'{sid}: no bound template page for {page["kind"]}')
            continue
        check_navigation(page, master)
        source_page = original.slides[master['slide_index'] - 1]
        roles = master.get('role_shapes', {})
        removed = set(master.get('remove_shape_ids', []))
        fixed = {shape.shape_id for shape in source_page.shapes} - set(roles.values()) - removed
        if (shared.native_fingerprint(source_page, included=fixed) != shared.native_fingerprint(slide, included=fixed)
            or master_signature(source_page) != master_signature(slide)):
            page_errors.append('fixed template objects, layout, master or theme changed')
        # Check removed sample placeholders by comparing the original XML, since new shapes may reuse their IDs.
        sample_xml = {str(shape._element.xml) for shape in source_page.shapes if shape.shape_id in removed}
        if sample_xml and any(str(shape._element.xml) in sample_xml for shape in slide.shapes):
            page_errors.append('template sample content remains')
        shapes = {shape.shape_id: shape for shape in slide.shapes}
        values = shared.role_values(page, master)
        for role, shape_id in roles.items():
            shape = shapes.get(shape_id)
            if shape is None or not shape.has_text_frame or shape.text != values[role]:
                page_errors.append('template text mismatch: ' + role)
        notes = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
        if notes is None or notes.text != shared.notes_text(page):
            page_errors.append('notes mismatch')
        body = [shape for shape in slide.shapes if shape.shape_id not in fixed and shape.shape_id not in roles.values()]
        actual_text = [text for shape in body for text in visible_text(shape)]
        actual_text.extend(shapes[shape_id].text for shape_id in roles.values()
                           if shape_id in shapes and shapes[shape_id].has_text_frame)
        for value in missing_text(page['on_slide'], actual_text):
            page_errors.append('planned editable text missing: ' + value)
        counts = Counter(kind for shape in body if (kind := object_kind(shape)) is not None)
        expected = Counter(obj['type'] for obj in page.get('objects', [])) | Counter(page.get('required_objects', {}))
        for kind, count in expected.items():
            if counts[kind] < count:
                page_errors.append(f'expected {count} editable/declared {kind} object(s), found {counts[kind]}')
        planned_charts = [
            (obj.get('chart_type', 'COLUMN_CLUSTERED'), [str(value) for value in obj['categories']],
             [(series['name'], tuple(series['values'])) for series in obj['series']])
            for obj in page.get('objects', []) if obj['type'] == 'chart'
        ]
        if planned_charts and planned_charts != [chart_data(shape.chart) for shape in body if shape.has_chart]:
            page_errors.append('editable chart types, categories or data differ from plan')
        errors.extend(f'{sid}: {error}' for error in page_errors)
        pages.append({'id': sid, 'objects': dict(counts), 'passed': not page_errors})
    return {
        'mode': 'native-editable',
        'pptx_sha256': shared.digest_file(pptx),
        'plan_sha256': shared.digest_file(path),
        'template': {'source_sha256': profile['source_sha256'], 'profile_sha256': shared.digest_file(profile_path),
                     'visual_status': profile.get('visual_status')},
        'font': plan.get('font'),
        'font_substitution': plan.get('font_substitution'),
        'target_app_acceptance': 'not-tested',
        'structure': {'passed': not errors, 'errors': errors, 'pages': pages},
        'visual_review': {'status': 'pending'},
    }


def accept(path, pptx, output, *, renders=None, reviewer=None, comment=None):
    if (reviewer is None) != (comment is None) or (reviewer is not None and not (reviewer.strip() and comment.strip())):
        raise ValueError('Provide a non-empty reviewer and comment together')
    if reviewer and renders is None:
        raise ValueError('Review existing final renders with --renders before recording visual acceptance')
    qa = output / 'qa.json'
    sheet = output / 'contact-sheet.jpg'
    if {path.resolve(), pptx.resolve()} & {qa.resolve(), sheet.resolve()}:
        raise ValueError('Acceptance output must be separate from its inputs')
    report = verify(path, pptx)
    shared.dump(qa, report)
    if not report['structure']['passed']:
        raise ValueError('\n'.join(report['structure']['errors']))
    if renders is None:
        renders = Path(tempfile.mkdtemp(prefix='renders-', dir=output))
        shared.render(SimpleNamespace(pptx=pptx, out=renders))
    proof = json.loads((renders / 'render.json').read_text(encoding='utf-8'))
    if proof['pptx_sha256'] != report['pptx_sha256']:
        raise ValueError('Render belongs to another PPT revision')
    images = {p.name: shared.digest_file(p) for p in renders.glob('*.png')}
    renderer = proof.get('renderer')
    if (not isinstance(renderer, str) or not renderer.strip() or not images
        or proof['pages'] != len(images) or len(images) != len(report['structure']['pages'])
        or images != proof['image_sha256']):
        raise ValueError('Render provenance, page count or image hashes do not match')
    shared.contact_sheet(SimpleNamespace(images=renders, out=sheet, columns=3))
    report['render_evidence'] = {
        **proof, 'directory': str(renders.resolve()), 'contact_sheet_sha256': shared.digest_file(sheet),
    }
    if reviewer:
        report['visual_review'] = {'status': 'passed', 'reviewer': reviewer, 'comment': comment}
    shared.dump(qa, report)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    for command in ('build', 'accept'):
        sub = commands.add_parser(command)
        sub.add_argument('--plan', required=True, type=Path)
        sub.add_argument('--out', required=True, type=Path)
        if command == 'accept':
            sub.add_argument('--pptx', required=True, type=Path)
            sub.add_argument('--renders', type=Path)
            sub.add_argument('--reviewer')
            sub.add_argument('--comment')
    args = parser.parse_args()
    try:
        if args.command == 'build':
            build(args.plan, args.out)
            print(args.out)
        else:
            report = accept(args.plan, args.pptx, args.out, renders=args.renders,
                            reviewer=args.reviewer, comment=args.comment)
            print(json.dumps({'qa': str(args.out / 'qa.json'), 'visual_review': report['visual_review']['status']}))
    except (BadZipFile, ValueError, OSError, KeyError, subprocess.SubprocessError) as error:
        parser.exit(1, str(error) + '\n')


if __name__ == '__main__':
    main()
