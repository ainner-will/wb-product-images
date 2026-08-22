from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
base=Path(r"H:\New project\wb-product-images\images")
out=base/"work_preview"; out.mkdir(exist_ok=True)
folders=sorted([p for p in base.iterdir() if p.is_dir() and p.name not in {'work_preview','outputs','node_modules'}],key=lambda p:p.name)
font=ImageFont.load_default(); tw,th=210,285
for page in range(0,len(folders),6):
 sel=folders[page:page+6]; rows=[]
 for folder in sel:
  ims=[]
  for i,p in enumerate(sorted([x for x in folder.iterdir() if x.suffix.lower() in {'.png','.jpg','.jpeg','.webp'}],key=lambda x:x.name)):
   im=Image.open(p).convert('RGB'); im.thumbnail((tw,th-30)); c=Image.new('RGB',(tw,th),'white'); c.paste(im,((tw-im.width)//2,25+(th-30-im.height)//2)); ImageDraw.Draw(c).text((4,4),f'{folder.name} #{i+1}',fill='black',font=font); ims.append(c)
  rows.append(ims)
 cols=max(map(len,rows)); sheet=Image.new('RGB',(cols*tw,len(rows)*th),(225,225,225))
 for r,row in enumerate(rows):
  for c,t in enumerate(row):sheet.paste(t,(c*tw,r*th))
 sheet.save(out/f'current_details_{page//6+1}.jpg',quality=92)
