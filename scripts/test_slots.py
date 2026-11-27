#!/usr/bin/env python3
"""Slot index contract tests. Fixtures are synthetic PPTX built in-process."""
import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace as NS

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt
from PIL import Image
from lxml import etree

SKILL = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL / 'scripts'))
import template_tools as tt
import skill_distribution as dist


class SlotIndexTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='slot-index-test-')
        self.base = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def deck(self, name, builders):
        prs = Presentation()
        prs.slide_width = Inches(13.333333)
        prs.slide_height = Inches(7.5)
        for build in builders:
            build(prs.slides.add_slide(prs.slide_layouts[6]))
        path = self.base / name
        prs.save(path)
        return path

    def test_single_text_shape_becomes_one_slot(self):
        holder = {}

        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
            holder['shape_id'] = box.shape_id

        slots = tt.index_template(self.deck('one.pptx', [build]))

        self.assertEqual(len(slots['pages']), 1)
        page = slots['pages'][0]
        self.assertEqual(page['slide_index'], 1)
        self.assertEqual(len(page['text_slots']), 1)
        slot = page['text_slots'][0]
        self.assertEqual(slot['current_text'], '模板标题')
        self.assertEqual(slot['address']['shape_id'], holder['shape_id'])

    def test_text_inside_group_is_indexed(self):
        holder = {}

        def build(slide):
            outer = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(.5))
            outer.text = '组外'
            inner = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(.5))
            inner.text = '组内'
            slide.shapes.add_group_shape([inner])
            holder['inner_id'] = inner.shape_id

        page = tt.index_template(self.deck('group.pptx', [build]))['pages'][0]

        self.assertEqual([s['current_text'] for s in page['text_slots']], ['组外', '组内'])
        self.assertIn(holder['inner_id'],
                      [s['address']['shape_id'] for s in page['text_slots']])

    def test_capacity_is_derived_from_box_and_font_size(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '容量测试'
            box.text_frame.word_wrap = True
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)

        slot = tt.index_template(self.deck('capacity.pptx', [build]))['pages'][0]['text_slots'][0]

        self.assertEqual(slot['capacity']['chars_per_line'], 8)
        self.assertEqual(slot['capacity']['max_lines'], 2)
        self.assertEqual(slot['capacity']['budget'], 19)

    def test_index_records_source_identity(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '来源'

        path = self.deck('meta.pptx', [build])
        slots = tt.index_template(path)

        self.assertEqual(slots['schema_version'], 1)
        self.assertEqual(slots['source_pptx'], 'meta.pptx')
        self.assertEqual(slots['source_sha256'],
                         hashlib.sha256(path.read_bytes()).hexdigest())
        self.assertEqual(slots['slide_size_cm'], [33.87, 19.05])

    def test_type_scale_ranks_distinct_font_sizes(self):
        def text(slide, label, size):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = label
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(size)

        def build(slide):
            text(slide, '大标题', 40)
            text(slide, '正文', 16)
            text(slide, '小标题', 24)

        slots = tt.index_template(self.deck('scale.pptx', [build]))

        self.assertEqual([t['size_pt'] for t in slots['type_scale']], [40, 24, 16])
        self.assertEqual([t['level'] for t in slots['type_scale']], [1, 2, 3])
        self.assertEqual(slots['type_scale'][0]['slot_count'], 1)
        by_text = {s['current_text']: s['level'] for s in slots['pages'][0]['text_slots']}
        self.assertEqual(by_text, {'大标题': 1, '正文': 3, '小标题': 2})

    def test_cli_index_writes_slots_file(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '落盘'

        path = self.deck('cli.pptx', [build])
        out = self.base / 'nested' / 'slots.json'

        tt.index(NS(pptx=path, out=out))

        written = json.loads(out.read_text(encoding='utf-8'))
        self.assertEqual(written['schema_version'], 1)
        self.assertEqual(written['source_pptx'], 'cli.pptx')
        self.assertEqual([s['current_text'] for s in written['pages'][0]['text_slots']], ['落盘'])

    def test_capacity_reports_shrink_to_fit(self):
        def build(slide):
            grow = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            grow.text = '自动缩字'
            grow.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            fixed = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
            fixed.text = '固定字号'

        by_text = {s['current_text']: s['capacity']
                   for s in tt.index_template(self.deck('autofit.pptx', [build]))['pages'][0]['text_slots']}

        self.assertTrue(by_text['自动缩字']['autofit'])
        self.assertFalse(by_text['固定字号']['autofit'])

    def test_slots_are_editable_until_an_agent_locks_them(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '01'

        slot = tt.index_template(self.deck('editable.pptx', [build]))['pages'][0]['text_slots'][0]

        self.assertIs(slot['editable'], True)

    def test_real_bound_template_addresses_resolve(self):
        source = SKILL / 'assets/master/source.pptx'
        if not source.is_file():
            self.skipTest('Real template not bound in this distribution')

        slots = tt.index_template(source)
        prs = Presentation(str(source))

        def resolve(shapes, shape_id):
            for shape in shapes:
                if shape.shape_id == shape_id:
                    return shape
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    found = resolve(shape.shapes, shape_id)
                    if found is not None:
                        return found
            return None

        checked = 0
        for page in slots['pages']:
            slide = prs.slides[page['slide_index'] - 1]
            for slot in page['text_slots']:
                shape = resolve(slide.shapes, slot['address']['shape_id'])
                self.assertIsNotNone(shape, slot['id'])
                para = shape.text_frame.paragraphs[slot['address']['paragraph']]
                self.assertEqual(''.join(run.text for run in para.runs),
                                 slot['current_text'], slot['id'])
                checked += 1

        self.assertGreater(checked, 0)
        self.assertLessEqual(len(slots['type_scale']), 6)

    def test_build_replaces_text_and_keeps_format(self):
        holder = {}

        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'
            run = box.text_frame.paragraphs[0].runs[0]
            run.font.size = Pt(32)
            run.font.bold = True
            holder['shape_id'] = box.shape_id

        template = self.deck('tpl.pptx', [build])
        slots = tt.index_template(template)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': '2026 年度复盘'}]}]}
        out = self.base / 'final.pptx'

        report = tt.build_deck(plan, slots, template, out)

        self.assertEqual(report['pages'], 1)
        prs = Presentation(str(out))
        self.assertEqual(len(prs.slides), 1)
        shape = next(s for s in prs.slides[0].shapes if s.shape_id == holder['shape_id'])
        run = shape.text_frame.paragraphs[0].runs[0]
        self.assertEqual(run.text, '2026 年度复盘')
        self.assertEqual(run.font.size.pt, 32)
        self.assertTrue(run.font.bold)

    def test_build_drops_unselected_template_pages(self):
        def page(text):
            def build(slide):
                box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
                box.text = text
            return build

        template = self.deck('three.pptx', [page('第一页'), page('第二页'), page('第三页')])
        slots = tt.index_template(template)
        slot_id = slots['pages'][1]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 2, 'edits': [{'slot': slot_id, 'text': '只留这页'}]}]}
        out = self.base / 'three-final.pptx'

        report = tt.build_deck(plan, slots, template, out)

        self.assertEqual(report['pages'], 1)
        prs = Presentation(str(out))
        self.assertEqual([s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame],
                         ['只留这页'])

    def test_build_can_reuse_one_template_page_twice(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '分章扉页'

        template = self.deck('reuse.pptx', [build])
        slots = tt.index_template(template)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': '第一章'}]},
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': '第二章'}]}]}
        out = self.base / 'reuse-final.pptx'

        tt.build_deck(plan, slots, template, out)

        prs = Presentation(str(out))
        self.assertEqual([next(s for s in slide.shapes if s.has_text_frame).text_frame.text
                          for slide in prs.slides], ['第一章', '第二章'])

    def test_build_warns_on_ellipsis_truncation_without_blocking(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'

        template = self.deck('ellipsis.pptx', [build])
        slots = tt.index_template(template)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': '这是一段被截断的说明...'}]}]}
        out = self.base / 'ellipsis-final.pptx'

        report = tt.build_deck(plan, slots, template, out)

        self.assertTrue(out.is_file())
        self.assertEqual(len(report['warnings']), 1)
        self.assertIn('省略号', report['warnings'][0])

    def test_build_rejects_placeholder_text(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'

        template = self.deck('placeholder.pptx', [build])
        slots = tt.index_template(template)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': 'Key Words Here'}]}]}
        out = self.base / 'placeholder-final.pptx'

        with self.assertRaisesRegex(ValueError, '占位'):
            tt.build_deck(plan, slots, template, out)

        self.assertFalse(out.exists())

    def test_build_rejects_unfilled_editable_slots(self):
        def build(slide):
            title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            title.text = '标题'
            subtitle = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
            subtitle.text = '副标题'

        template = self.deck('unfilled.pptx', [build])
        slots = tt.index_template(template)
        title_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': title_id, 'text': '新标题'}]}]}

        with self.assertRaisesRegex(ValueError, '未覆盖'):
            tt.build_deck(plan, slots, template, self.base / 'unfilled-final.pptx')

    def test_build_rejects_slots_index_from_another_template(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'

        template = self.deck('source-a.pptx', [build])

        def other_build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
            box.text = '另一个模板标题'

        other = self.deck('source-b.pptx', [other_build])
        slots = tt.index_template(template)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': '换字'}]}]}

        with self.assertRaisesRegex(ValueError, '来源'):
            tt.build_deck(plan, slots, other, self.base / 'mismatch-final.pptx')

    def test_build_rejects_unknown_slot_id(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'

        template = self.deck('unknown-slot.pptx', [build])
        slots = tt.index_template(template)
        real_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': real_id, 'text': '正文'},
                                          {'slot': 's1_sh999_p0', 'text': '多余'}]}]}

        with self.assertRaisesRegex(ValueError, '槽位'):
            tt.build_deck(plan, slots, template, self.base / 'unknown-final.pptx')

    def test_build_rejects_out_of_range_source_slide(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'

        template = self.deck('range.pptx', [build])
        slots = tt.index_template(template)
        real_id = slots['pages'][0]['text_slots'][0]['id']
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 9, 'edits': [{'slot': real_id, 'text': '正文'}]}]}

        with self.assertRaisesRegex(ValueError, '页'):
            tt.build_deck(plan, slots, template, self.base / 'range-final.pptx')

    def test_cli_build_writes_deck_next_to_slots(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'

        template = self.deck('cli-build.pptx', [build])
        slots = tt.index_template(template)
        master = self.base / 'master'
        master.mkdir()
        (master / 'cli-build.pptx').write_bytes(template.read_bytes())
        slots_path = master / 'slots.json'
        tt.dump(slots_path, slots)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        plan_path = self.base / 'plan.json'
        tt.dump(plan_path, {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': '端到端标题'}]}]})
        out = self.base / 'out' / 'final.pptx'

        tt.build(NS(plan=plan_path, slots=slots_path, out=out, template=None))

        self.assertTrue(out.is_file())
        prs = Presentation(str(out))
        self.assertEqual([s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame],
                         ['端到端标题'])

    def test_build_works_against_real_template(self):
        source = SKILL / 'assets/master/source.pptx'
        if not source.is_file():
            self.skipTest('Real template not bound in this distribution')

        slots = tt.index_template(source)
        page = slots['pages'][0]
        planned = [f'第{i}个槽' for i in range(1, len(page['text_slots']) + 1)]
        plan = {'schema_version': 1, 'slides': [{
            'source_slide': 1,
            'edits': [{'slot': slot['id'], 'text': text}
                      for slot, text in zip(page['text_slots'], planned)]}]}
        out = self.base / 'real-final.pptx'

        report = tt.build_deck(plan, slots, source, out)

        self.assertEqual(report['pages'], 1)
        prs = Presentation(str(out))
        visible = [s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame]
        paragraphs = '\n'.join(visible).split('\n')
        for text in planned:
            self.assertIn(text, paragraphs)

    def test_build_warns_when_text_exceeds_capacity(self):
        def build(slide):
            box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            box.text = '模板标题'
            box.text_frame.word_wrap = True
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(32)

        template = self.deck('overflow.pptx', [build])
        slots = tt.index_template(template)
        slot_id = slots['pages'][0]['text_slots'][0]['id']
        too_long = '这一段文字明显超过了文本框容量应该触发提示' * 2
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': 1, 'edits': [{'slot': slot_id, 'text': too_long}]}]}
        out = self.base / 'overflow-final.pptx'

        report = tt.build_deck(plan, slots, template, out)

        self.assertTrue(out.is_file())
        self.assertEqual(len(report['warnings']), 1)
        self.assertIn('偏长', report['warnings'][0])

    def built(self, name='accept', texts=('标题', '副标题')):
        """Build a two-page deck from a template with one fixed picture and two slots."""
        asset = self.base / 'fixed.png'
        Image.new('RGB', (100, 40), '#11298C').save(asset)

        def build(slide):
            slide.shapes.add_picture(str(asset), Inches(.1), Inches(.1),
                                    width=Inches(.8), height=Inches(.3))
            for row, text in enumerate(texts):
                box = slide.shapes.add_textbox(Inches(1), Inches(1 + row), Inches(4), Inches(1))
                box.text = text
                box.text_frame.paragraphs[0].runs[0].font.size = Pt(24)

        template = self.deck(f'{name}-tpl.pptx', [build, build])
        slots = tt.index_template(template)
        plan = {'schema_version': 1, 'slides': [
            {'source_slide': index,
             'edits': [{'slot': slot['id'], 'text': f'第{index}页第{n}槽'}
                       for n, slot in enumerate(page['text_slots'], 1)]}
            for index, page in enumerate(slots['pages'], 1)]}
        plan_path = self.base / f'{name}-plan.json'
        tt.dump(plan_path, plan)
        pptx = self.base / f'{name}-final.pptx'
        tt.build_deck(plan, slots, template, pptx)
        return template, slots, plan_path, pptx

    def test_verify_accepts_slot_inside_group(self):
        def build(slide):
            outer = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            outer.text = '组外'
            inner = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
            inner.text = '组内'
            slide.shapes.add_group_shape([inner])

        template = self.deck('group-verify-tpl.pptx', [build])
        slots = tt.index_template(template)
        plan = {'schema_version': 1, 'slides': [{
            'source_slide': 1,
            'edits': [{'slot': slot['id'], 'text': f'新{n}'}
                      for n, slot in enumerate(slots['pages'][0]['text_slots'], 1)]}]}
        pptx = self.base / 'group-verify-final.pptx'
        tt.build_deck(plan, slots, template, pptx)

        report = tt.verify_deck(plan, slots, template, pptx)

        self.assertTrue(report['structure']['passed'], report['structure']['errors'])

    def test_verify_passes_on_built_deck(self):
        template, slots, plan, pptx = self.built()

        report = tt.verify_deck(plan, slots, template, pptx)

        self.assertTrue(report['structure']['passed'], report['structure']['errors'])
        self.assertEqual(len(report['structure']['pages']), 2)
        self.assertEqual(report['mode'], 'page-slot')

    def test_verify_rejects_placeholder_left_in_output(self):
        template, slots, plan, pptx = self.built('residue')
        prs = Presentation(str(pptx))
        target = next(s for s in prs.slides[0].shapes if s.has_text_frame)
        target.text_frame.paragraphs[0].runs[0].text = 'Key Words Here'
        dirty = self.base / 'residue.pptx'
        prs.save(str(dirty))

        report = tt.verify_deck(plan, slots, template, dirty)

        self.assertFalse(report['structure']['passed'])
        self.assertTrue(any('占位' in error for error in report['structure']['errors']))

    def test_verify_rejects_moved_fixed_object(self):
        template, slots, plan, pptx = self.built('tamper')
        prs = Presentation(str(pptx))
        picture = next(s for s in prs.slides[1].shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
        picture.left += Inches(.2)
        moved = self.base / 'tamper-moved.pptx'
        prs.save(str(moved))

        report = tt.verify_deck(plan, slots, template, moved)

        self.assertFalse(report['structure']['passed'])
        self.assertIn('fixed template objects changed', report['structure']['errors'][0])

    def test_verify_rejects_text_drift(self):
        template, slots, plan, pptx = self.built('drift')
        prs = Presentation(str(pptx))
        target = json.loads(Path(plan).read_text(encoding='utf-8'))['slides'][0]['edits'][0]
        address = next(slot['address'] for page in slots['pages']
                       for slot in page['text_slots'] if slot['id'] == target['slot'])
        shape = next(s for s in prs.slides[0].shapes if s.shape_id == address['shape_id'])
        shape.text_frame.paragraphs[address['paragraph']].runs[0].text = '被外部改过'
        drifted = self.base / 'drift-changed.pptx'
        prs.save(str(drifted))

        report = tt.verify_deck(plan, slots, template, drifted)

        self.assertFalse(report['structure']['passed'])
        self.assertIn('slot text mismatch', report['structure']['errors'][0])

    def renders_for(self, pptx, name='renders'):
        """Stand in for the host renderer: one PNG per slide plus a render proof."""
        import deck_pipeline as dp
        folder = self.base / name
        folder.mkdir()
        pages = len(Presentation(str(pptx)).slides)
        for index in range(1, pages + 1):
            Image.new('RGB', (1920, 1080), '#eef1f7').save(folder / f'slide-{index}.png')
        images = {p.name: dp.digest_file(p) for p in folder.glob('slide-*.png')}
        dp.dump(folder / 'render.json', {'pptx_sha256': dp.digest_file(pptx), 'pages': pages,
                                        'renderer': 'test-fixture',
                                        'image_sha256': images})
        return folder

    def test_accept_records_structure_and_visual_review(self):
        template, slots, plan, pptx = self.built('accept-ok')
        work = self.base / 'accept-ok-work'
        work.mkdir()
        renders = self.renders_for(pptx)

        report = tt.accept_deck(plan, slots, template, pptx, work, renders=renders,
                                reviewer='agent', comment='合成夹具，仅验证验收链路')

        qa = json.loads((work / 'qa.json').read_text(encoding='utf-8'))
        self.assertTrue(qa['structure']['passed'])
        self.assertEqual(qa['visual_review']['status'], 'passed')
        self.assertEqual(qa['visual_review']['comment'], '合成夹具，仅验证验收链路')
        self.assertTrue((work / 'contact-sheet.jpg').is_file())
        self.assertEqual(report['render_evidence']['pages'], 2)

    def test_accept_rejects_render_from_another_revision(self):
        template, slots, plan, pptx = self.built('accept-stale')
        work = self.base / 'accept-stale-work'
        work.mkdir()
        renders = self.renders_for(pptx)
        revised = self.base / 'accept-stale-revised.pptx'
        revised.write_bytes(pptx.read_bytes() + b'')  # same structure, new bytes
        prs = Presentation(str(revised))
        prs.core_properties.title = '重新导出的同一套版式'
        prs.save(str(revised))

        with self.assertRaisesRegex(ValueError, 'another revision'):
            tt.accept_deck(plan, slots, template, revised, work, renders=renders,
                            reviewer='agent', comment='渲染属于旧版本，应被拒绝')

    def test_accept_without_reviewer_keeps_visual_pending(self):
        template, slots, plan, pptx = self.built('accept-pending')
        work = self.base / 'accept-pending-work'
        work.mkdir()
        renders = self.renders_for(pptx)

        tt.accept_deck(plan, slots, template, pptx, work, renders=renders)

        qa = json.loads((work / 'qa.json').read_text(encoding='utf-8'))
        self.assertEqual(qa['visual_review']['status'], 'pending')
        self.assertIn('render_evidence', qa)

    def test_accept_requires_reviewer_and_comment_together(self):
        template, slots, plan, pptx = self.built('accept-pair')
        work = self.base / 'accept-pair-work'
        work.mkdir()
        renders = self.renders_for(pptx)

        with self.assertRaisesRegex(ValueError, 'together'):
            tt.accept_deck(plan, slots, template, pptx, work, renders=renders,
                            reviewer='agent', comment='   ')


class FontInheritanceTests(unittest.TestCase):
    """Placeholder text with no explicit size must resolve through the OOXML chain."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='font-chain-test-')
        self.base = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def save(self, prs, name):
        path = self.base / name
        prs.save(path)
        return path

    def test_title_placeholder_inherits_master_title_style(self):
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = '继承标题'
        path = self.save(prs, 'title.pptx')

        slot = tt.index_template(path)['pages'][0]['text_slots'][0]

        self.assertEqual(slot['current_text'], '继承标题')
        self.assertEqual(slot['capacity']['font_size_pt'], 44.0)

    def test_layout_placeholder_style_beats_master_tx_style(self):
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
        for ph in prs.slide_layouts[1].placeholders:
            if ph.placeholder_format.idx != 0:
                continue
            body = ph._element.find(tt.P_NS + 'txBody')
            lst = body.find(f'{tt.A_NS}lstStyle')
            lst.append(etree.fromstring(
                f'<a:lvl1pPr xmlns:a="{tt.A_NS[1:-1]}"><a:defRPr sz="3200"/></a:lvl1pPr>'))
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = '布局覆盖标题'
        path = self.save(prs, 'layout.pptx')

        slot = tt.index_template(path)['pages'][0]['text_slots'][0]

        self.assertEqual(slot['capacity']['font_size_pt'], 32.0)

    def test_plain_textbox_without_size_keeps_default(self):
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
        box = prs.slides.add_slide(prs.slide_layouts[6]).shapes.add_textbox(
            Inches(1), Inches(1), Inches(4), Inches(1))
        box.text = '无字号文本框'
        path = self.save(prs, 'plain.pptx')

        slot = tt.index_template(path)['pages'][0]['text_slots'][0]

        self.assertEqual(slot['capacity']['font_size_pt'], tt.DEFAULT_SIZE_PT)

    def test_vertical_placeholder_types_map_to_their_buckets(self):
        for name, bucket in [('VERTICAL_TITLE', 'titleStyle'),
                             ('VERTICAL_BODY', 'bodyStyle'),
                             ('VERTICAL_OBJECT', 'bodyStyle'),
                             ('TITLE', 'titleStyle'), ('CENTER_TITLE', 'titleStyle'),
                             ('BODY', 'bodyStyle'), ('SUBTITLE', 'bodyStyle'),
                             ('OBJECT', 'bodyStyle')]:
            self.assertEqual(tt.STYLE_BUCKETS.get(name), bucket, name)


class DistributionTests(unittest.TestCase):
    """A skill release ships a per-file manifest and an incremental update index."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='dist-test-')
        self.base = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def tree(self, name, files):
        root = self.base / name
        for path, content in files.items():
            target = root / path
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(content if isinstance(content, bytes) else content.encode())
        return root

    def test_manifest_matches_the_published_package_contents(self):
        root = self.tree('pkg', {'.gitignore': 'x', '.DS_Store': 'junk',
                                 '.github/workflows/ci.yml': 'y',
                                 'SKILL.md': 'v', 'scripts/a.py': 'code'})

        files = sorted(dist.shipped_files(root))

        self.assertEqual(files, ['SKILL.md', 'scripts/a.py'])

    def test_delta_lists_added_modified_and_removed(self):
        old = {'a.md': 'same', 'b.md': 'before', 'gone.md': 'x'}
        new = {'a.md': 'same', 'b.md': 'after', 'new.md': 'y'}

        delta = dist.delta(old, new)

        self.assertEqual(delta['added'], ['new.md'])
        self.assertEqual(delta['modified'], ['b.md'])
        self.assertEqual(delta['removed'], ['gone.md'])

    def test_manifest_records_hash_and_size_per_file(self):
        root = self.tree('skill', {'SKILL.md': 'v1', 'scripts/a.py': 'print(1)',
                                   '.venv/lib/x.pyc': 'cache', 'scripts/__pycache__/a.cpython.pyc': 'c'})

        manifest = dist.build_manifest(root)

        self.assertEqual(manifest['schema'], dist.MANIFEST_SCHEMA)
        self.assertEqual(sorted(manifest['files']), ['SKILL.md', 'scripts/a.py'])
        entry = manifest['files']['SKILL.md']
        self.assertEqual(entry['sha256'], hashlib.sha256(b'v1').hexdigest())
        self.assertEqual(entry['size'], 2)

    def test_history_appends_and_replaces_versions(self):
        index = dist.new_index('https://example.com/skill')

        dist.record_release(index, '0.3.0', {'added': ['a.md'], 'modified': [], 'removed': []},
                            'first')
        dist.record_release(index, '0.3.1', {'added': [], 'modified': ['b.md'], 'removed': []},
                            'second')
        dist.record_release(index, '0.3.1', {'added': ['c.md'], 'modified': [], 'removed': []},
                            'rebuilt')

        self.assertEqual(index['schema'], dist.SCHEMA)
        self.assertEqual(index['latest_version'], '0.3.1')
        self.assertEqual([h['version'] for h in index['history']], ['0.3.0', '0.3.1'])
        self.assertEqual(index['history'][1]['added'], ['c.md'])
        self.assertEqual(index['history'][1]['summary'], 'rebuilt')

    def test_plan_update_unions_versions_newer_than_local(self):
        index = dist.new_index('https://example.com/skill')
        dist.record_release(index, '0.2.0', {'added': ['old.md'], 'modified': [], 'removed': []}, 'x')
        dist.record_release(index, '0.3.0', {'added': ['a.md', 'temp.md'], 'modified': ['b.md'],
                                             'removed': []}, 'x')
        dist.record_release(index, '0.3.1', {'added': [], 'modified': ['c.md'],
                                             'removed': ['old.md', 'temp.md']}, 'x')

        plan = dist.plan_update(index, '0.2.0')

        self.assertTrue(plan['needs_update'])
        self.assertEqual(plan['target_version'], '0.3.1')
        self.assertEqual(plan['added'], ['a.md'])
        self.assertEqual(plan['modified'], ['b.md', 'c.md'])
        self.assertEqual(plan['removed'], ['old.md', 'temp.md'])

    def test_plan_update_is_noop_when_current(self):
        index = dist.new_index('https://example.com/skill')
        dist.record_release(index, '0.3.1', {'added': ['a.md'], 'modified': [], 'removed': []}, 'x')

        plan = dist.plan_update(index, '0.3.1')

        self.assertFalse(plan['needs_update'])
        self.assertEqual((plan['added'], plan['modified'], plan['removed']), ([], [], []))

    def test_apply_update_writes_verified_files_and_prunes_removed(self):
        target = self.tree('installed', {'keep.md': 'old', 'stale.md': 'bye', 'VERSION': '0.3.0'})
        remote = {'keep.md': b'new', 'fresh.md': b'hello', 'VERSION': b'0.3.1'}
        remote_manifest = {'files': {name: {'sha256': hashlib.sha256(data).hexdigest()}
                                     for name, data in remote.items()}}

        report = dist.apply_update(
            target, {'added': ['fresh.md'], 'modified': ['keep.md', 'VERSION'],
                     'removed': ['stale.md']},
            lambda name: remote[name], remote_manifest, '0.3.1')

        self.assertEqual(report['version'], '0.3.1')
        self.assertEqual((target / 'keep.md').read_bytes(), b'new')
        self.assertEqual((target / 'fresh.md').read_bytes(), b'hello')
        self.assertFalse((target / 'stale.md').exists())
        self.assertEqual((target / 'VERSION').read_text().strip(), '0.3.1')
        for name, data in remote.items():
            self.assertEqual((target / name).read_bytes(), data)

    def test_apply_update_requires_every_file_in_the_manifest(self):
        target = self.tree('noentry', {'keep.md': 'old'})

        with self.assertRaisesRegex(ValueError, '缺少'):
            dist.apply_update(target, {'added': ['ghost.md'], 'modified': [], 'removed': []},
                              lambda name: b'anything', {'files': {}}, '0.4.0')

        self.assertFalse((target / 'ghost.md').exists())

    def test_re_added_file_is_fetched_not_skipped(self):
        index = dist.new_index('https://example.com/skill')
        dist.record_release(index, '0.2.0', {'added': ['a.md'], 'modified': [], 'removed': []}, 'x')
        dist.record_release(index, '0.3.0', {'added': [], 'modified': [], 'removed': ['a.md']}, 'x')
        dist.record_release(index, '0.3.1', {'added': ['a.md'], 'modified': [], 'removed': []}, 'x')

        plan = dist.plan_update(index, '0.2.0')

        self.assertIn('a.md', plan['added'])

    def test_apply_update_rejects_hash_mismatch_without_partial_writes(self):
        target = self.tree('integrity', {'keep.md': 'old'})
        remote_manifest = {'files': {'keep.md': {'sha256': '0' * 64}}}

        with self.assertRaisesRegex(ValueError, 'sha256'):
            dist.apply_update(target, {'added': [], 'modified': ['keep.md'], 'removed': []},
                              lambda name: b'tampered', remote_manifest, '0.4.0')

        self.assertEqual((target / 'keep.md').read_text(), 'old')
        self.assertFalse((target / 'VERSION').exists())

    def test_cli_build_writes_manifest_and_records_delta(self):
        src = self.tree('src', {'SKILL.md': 'v2', 'scripts/s.py': 'code',
                                 '.venv/junk.py': 'x'})
        out = self.base / 'mirror'
        out.mkdir()
        previous = {'files': {'SKILL.md': {'sha256': hashlib.sha256(b'v1').hexdigest()},
                              'old.md': {'sha256': 'a' * 64}}}
        dist.dump(out / 'manifest.json', previous)
        index = dist.new_index('https://example.com/skill')

        dist.build_release(src, out, '0.3.0', '增量分发', index)

        manifest = json.loads((out / 'manifest.json').read_text(encoding='utf-8'))
        self.assertEqual(sorted(manifest['files']), ['SKILL.md', 'scripts/s.py'])
        index = json.loads((out / 'updates.json').read_text(encoding='utf-8'))
        self.assertEqual(index['latest_version'], '0.3.0')
        self.assertEqual(index['history'][0]['added'], ['scripts/s.py'])
        self.assertEqual(index['history'][0]['modified'], ['SKILL.md'])
        self.assertEqual(index['history'][0]['removed'], ['old.md'])


if __name__ == '__main__':
    unittest.main(verbosity=2)
