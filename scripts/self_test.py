import sys, json, copy, unittest, shutil, tempfile
from pathlib import Path
from types import SimpleNamespace as NS
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
SKILL=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(SKILL/'scripts'))
import deck_pipeline as dp
import template_tools as tt
TEMP=tempfile.TemporaryDirectory(prefix='enterprise-ppt-self-test-')
BASE=Path(TEMP.name)

def sample_image(path,n):
 im=Image.new('RGB',(1920,1080),(240,245,250))
 draw=ImageDraw.Draw(im);draw.rectangle((120,300,1700,890),outline='#11298C',width=5);draw.text((200,400),f'Synthetic pipeline fixture {n}',fill='#11298C')
 im.save(path)

def fixture(name,mode='full-page-image',profile=None):
 folder=BASE/name;folder.mkdir(parents=True,exist_ok=True)
 d=json.loads((SKILL/'assets/deck-example.json').read_text());d['mode']=mode
 d['require_exact_model']=False  # Synthetic fixtures have no image backend.
 d['image_provider']='test-fixture'
 if mode == 'locked-template' and not profile: d['template_profile']=str(folder/'unbound/profile.json')
 if profile:d['template_profile']=str(profile)
 for i in range(1,3):sample_image(folder/f'P{i:03d}.png',i)
 plan=folder/'deck.json';dp.dump(plan,d)
 return folder,plan

def reviewed(plan):
 for i in range(1,3):
  dp.record(NS(plan=plan,slide=f'P{i:03d}',image=plan.parent/f'P{i:03d}.png',actual_model='synthetic-test-fixture',tool='test-fixture-builder'))
  dp.approve(NS(plan=plan,slide=f'P{i:03d}',reviewer='test-fixture',checks='text,numbers,diagram,master,crop',comment='Synthetic fixture; validates pipeline only, no model generation claimed'))

def make_template(folder):
 if (folder/'bound').exists():shutil.rmtree(folder/'bound')
 prs=Presentation();prs.slide_width=Inches(13.333333);prs.slide_height=Inches(7.5)
 masters={}
 for i,kind in enumerate(['cover','body'],1):
  s=prs.slides.add_slide(prs.slide_layouts[6])
  # Fixed native image exercises relationship ID cloning.
  asset=folder/'fixed.png';Image.new('RGB',(100,40),'#11298C').save(asset)
  s.shapes.add_picture(str(asset), Inches(.1),Inches(.1),width=Inches(.8),height=Inches(.3))
  roles={}
  items=[('title','Template title',.7,.2,10,.5),('subtitle','Template subtitle',.7,1,10,.4)]
  if kind=='body':items+=[('page_number','4',12.4,7.1,.5,.25)]
  else:items+=[('footer','Template organization',1,5.8,10,.6)]
  for role,text,x,y,w,h in items:
   sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));sh.text=text
   sh.text_frame.paragraphs[0].runs[0].font.size=Pt(24)
   sh.text_frame.paragraphs[0].runs[0].font.color.rgb=RGBColor.from_string('11298C');roles[role]=sh.shape_id
  sh=s.shapes.add_textbox(Inches(2),Inches(3),Inches(8),Inches(1));sh.text='Sample content to remove'
  masters[kind]={'slide_index':i,'role_shapes':roles,'remove_shape_ids':[sh.shape_id]}
  if kind=='body':masters[kind]['content_box']=[.055,.235,.9,.67]
 src=folder/'synthetic-template.pptx';prs.save(src)
 cfg=folder/'binding.json';dp.dump(cfg,{'masters':masters,'provenance':'SYNTHETIC TEST TEMPLATE, not user original'})
 tt.bind(NS(pptx=src,config=cfg,out=folder/'bound',replace=False))
 # Structural test fixture deliberately does not claim a real visual review.
 pp=folder/'bound/profile.json';p=json.loads(pp.read_text());p['visual_status']='verified';p['provenance']='SYNTHETIC TEST ONLY';dp.dump(pp,p)
 return pp

class PipelineTests(unittest.TestCase):
 def test_version_contract(self):
  import re, subprocess
  version=(SKILL/'VERSION').read_text().strip()
  self.assertRegex(version,r'^\d+\.\d+\.\d+$')
  metadata=re.search(r'^  version: "([^"\n]+)"$',(SKILL/'SKILL.md').read_text(),re.M)
  self.assertIsNotNone(metadata);self.assertEqual(metadata.group(1),version)
  self.assertIn('## '+version+' (',(SKILL/'CHANGELOG.md').read_text())
  result=subprocess.run([sys.executable,str(SKILL/'scripts/deck_pipeline.py'),'--version'],capture_output=True,text=True)
  self.assertEqual(result.returncode,0);self.assertEqual(result.stdout.strip(),'magic-ppt '+version)
 def test_01_prepare(self):
  folder,plan=fixture('prepare')
  dp.prepare(NS(plan=plan,out=folder/'build'))
  jobs=json.loads((folder/'build/jobs.json').read_text());self.assertEqual(len(jobs['jobs']),2)
  self.assertIn('主标题：',(folder/'build/prompts/P002.txt').read_text())
  self.assertIn('禁止',(folder/'build/prompts/P002.txt').read_text())
 def test_02_full_page_assembly(self):
  folder,plan=fixture('full');reviewed(plan)
  dp.assemble(NS(plan=plan,out=folder/'final.pptx'))
  prs=Presentation(folder/'final.pptx');d=json.loads(plan.read_text())
  self.assertEqual(len(prs.slides),2);self.assertEqual(prs.slides[1].notes_slide.notes_text_frame.text,dp.notes_text(d['slides'][1]))
  self.assertEqual(len(prs.slides[0].shapes),1)
 def test_03_reject_bad_ratio(self):
  folder,plan=fixture('bad-ratio');Image.new('RGB',(1920,1000)).save(folder/'wrong.png')
  with self.assertRaisesRegex(ValueError,'16:9'):dp.record(NS(plan=plan,slide='P001',image=folder/'wrong.png',actual_model='test',tool='test'))
 def test_04_reject_stale_image(self):
  folder,plan=fixture('stale-image');reviewed(plan);Image.new('RGB',(1920,1080),'red').save(folder/'P001.png')
  with self.assertRaisesRegex(ValueError,'image changed'):dp.assemble(NS(plan=plan,out=folder/'final.pptx'))
 def test_05_reject_stale_content(self):
  folder,plan=fixture('stale-content');reviewed(plan);d=json.loads(plan.read_text());d['slides'][1]['title']='Changed after QA';dp.dump(plan,d)
  with self.assertRaisesRegex(ValueError,'changed after review'):dp.assemble(NS(plan=plan,out=folder/'final.pptx'))
 def test_06_reject_unknown_exact_model(self):
  folder,plan=fixture('exact-model');d=json.loads(plan.read_text());d['require_exact_model']=True;dp.dump(plan,d)
  dp.record(NS(plan=plan,slide='P001',image=folder/'P001.png',actual_model='backend-not-exposed',tool='test'))
  with self.assertRaisesRegex(ValueError,'not confirmed'):dp.approve(NS(plan=plan,slide='P001',reviewer='test',checks='text,numbers,diagram,master,crop',comment='test'))
 def test_07_reject_missing_original(self):
  folder,plan=fixture('missing-template','locked-template');reviewed(plan)
  with self.assertRaisesRegex(ValueError,'not bound'):dp.assemble(NS(plan=plan,out=folder/'final.pptx'))
 def test_08_reject_duplicate_pages(self):
  folder,plan=fixture('duplicates');(folder/'P002.png').write_bytes((folder/'P001.png').read_bytes());reviewed(plan)
  with self.assertRaisesRegex(ValueError,'Duplicate'):dp.assemble(NS(plan=plan,out=folder/'final.pptx'))
 def test_09_locked_template_assembly_and_tamper(self):
  folder=BASE/'locked';folder.mkdir(exist_ok=True);pp=make_template(folder)
  folder,plan=fixture('locked','locked-template',pp);reviewed(plan)
  dp.assemble(NS(plan=plan,out=folder/'final.pptx'))
  prs=Presentation(folder/'final.pptx');d=json.loads(plan.read_text());p=json.loads(pp.read_text())
  sid=p['masters']['body']['role_shapes']['title']
  self.assertEqual(next(x for x in prs.slides[1].shapes if x.shape_id==sid).text,d['slides'][1]['title'])
  pic=next(x for x in prs.slides[1].shapes if x.name=='GeneratedContent');self.assertAlmostEqual(pic.crop_top,.235)
  self.assertNotIn('Sample content', '\n'.join(x.text for x in prs.slides[1].shapes if x.has_text_frame))
  # Move fixed native logo after export: verification must reject.
  prs.slides[1].shapes[0].left+=Inches(.2);prs.save(folder/'tampered.pptx')
  with self.assertRaisesRegex(ValueError,'fixed master'):dp.verify_deck(d,plan,folder/'tampered.pptx')
 def test_10_retry_limit(self):
  folder,plan=fixture('retry')
  for _ in range(3):dp.record(NS(plan=plan,slide='P001',image=folder/'P001.png',actual_model='test',tool='test'))
  with self.assertRaisesRegex(ValueError,'limit'):dp.record(NS(plan=plan,slide='P001',image=folder/'P001.png',actual_model='test',tool='test'))
 def test_11_role_overlap(self):
  folder=BASE/'overlap';folder.mkdir(exist_ok=True);pp=make_template(folder)
  p=json.loads(pp.read_text());p['masters']['body']['content_box']=[0,0,1,1]
  with self.assertRaisesRegex(ValueError,'overlaps'):tt.check_config(p,Presentation(pp.parent/'source.pptx'))
 def test_12_reject_low_resolution(self):
  folder,plan=fixture('low-res');Image.new('RGB',(1280,720)).save(folder/'low.png')
  with self.assertRaisesRegex(ValueError,'below'):dp.record(NS(plan=plan,slide='P001',image=folder/'low.png',actual_model='test',tool='test'))
 def test_13_current_chatgpt_tool_without_backend_name(self):
  folder,plan=fixture('current-chatgpt');d=json.loads(plan.read_text());d['image_provider']='chatgpt-built-in';d.pop('requested_image_model');d.pop('require_exact_model');dp.dump(plan,d)
  dp.record(NS(plan=plan,slide='P001',image=folder/'P001.png',actual_model='backend-not-exposed',tool='image_gen.imagegen'))
  dp.approve(NS(plan=plan,slide='P001',reviewer='test',checks='text,numbers,diagram,master,crop',comment='Synthetic provenance contract; no real image generation claimed'))
  _,loaded=dp.load_plan(plan);self.assertFalse(loaded['require_exact_model'])
  with self.assertRaisesRegex(ValueError,'built-in'):dp.model_check(loaded,{'tool':'third-party','actual_model':'backend-not-exposed'})
 def test_host_tool_provenance_and_invalidation(self):
  folder,plan=fixture('host-tool');d=json.loads(plan.read_text());d['image_provider']='host-tool';d['image_tool']='configured-host-image';d.pop('requested_image_model',None);dp.dump(plan,d)
  _,loaded=dp.load_plan(plan);self.assertEqual(loaded['requested_image_model'],'Current configured host image tool')
  dp.record(NS(plan=plan,slide='P001',image=folder/'P001.png',actual_model='backend-not-exposed',tool='configured-host-image'))
  dp.approve(NS(plan=plan,slide='P001',reviewer='test',checks='text,numbers,diagram,master,crop',comment='Synthetic host-tool provenance fixture'))
  with self.assertRaisesRegex(ValueError,'mismatch'):dp.model_check(loaded,{'tool':'other-tool','actual_model':'backend-not-exposed'})
  _,approved=dp.load_plan(plan);changed=dict(approved);changed['image_tool']='other-tool'
  self.assertNotEqual(dp.page_digest(approved,approved['slides'][0],plan),dp.page_digest(changed,changed['slides'][0],plan))
  changed.pop('image_tool');dp.dump(plan,changed)
  with self.assertRaisesRegex(ValueError,'requires image_tool'):dp.load_plan(plan)
  changed['image_provider']='unknown-provider';dp.dump(plan,changed)
  with self.assertRaisesRegex(ValueError,'Unsupported'):dp.load_plan(plan)
 def test_14_real_template_native_roles_and_watermark(self):
  real=SKILL/'assets/master/profile.json'
  if not real.is_file():self.skipTest('Real template not bound in this distribution')
  profile=json.loads(real.read_text());tt.check_config(profile,Presentation(real.parent/'source.pptx'))
  self.assertEqual(profile['masters']['body']['page_number_mode'],'inherited-field')
  self.assertEqual(dp.role_values({'title':'Title','subtitle':'Subtitle','order':1,'on_slide':[]},profile['masters']['cover'])['title'],'Title\nSubtitle')
  folder,plan=fixture('real-alpha','locked-template',real);_,d=dp.load_plan(plan);s=d['slides'][1]
  with self.assertRaisesRegex(ValueError,'Transparent'):dp.transparency_check(d,s,plan,folder/'P002.png')
  im=Image.new('RGBA',(1920,1080),(0,0,0,0));draw=ImageDraw.Draw(im);draw.rectangle((110,300,900,650),fill=(17,41,140,255));im.save(folder/'alpha.png')
  dp.transparency_check(d,s,plan,folder/'alpha.png')
  draw.rectangle((1580,800,1700,900),fill=(255,255,255,255));im.save(folder/'covered.png')
  with self.assertRaisesRegex(ValueError,'protected'):dp.transparency_check(d,s,plan,folder/'covered.png')

if __name__=='__main__':
 try:unittest.main(verbosity=2)
 finally:TEMP.cleanup()
