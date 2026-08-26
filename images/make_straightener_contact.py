from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

root = Path('.')
out = root / 'straightener_preview' / 'contacts'
out.mkdir(parents=True, exist_ok=True)
font = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 20)
files = []
for folder in sorted([p for p in root.iterdir() if p.is_dir() and p.name not in {'node_modules','straightener_preview','outputs'}]):
    imgs = sorted([p for p in folder.iterdir() if p.suffix.lower() in {'.png','.jpg','.jpeg','.webp'}], key=lambda p:p.name)
    files.append((folder.name, imgs))
for folder, imgs in files:
    tw, th, header = 360, 270, 34
    cols = 4
    rows = (len(imgs)+cols-1)//cols
    canvas = Image.new('RGB', (cols*tw, rows*(th+header)), 'white')
    d = ImageDraw.Draw(canvas)
    for i,p in enumerate(imgs):
        im = Image.open(p).convert('RGB')
        im.thumbnail((tw-10, th-10))
        x = (i%cols)*tw + (tw-im.width)//2
        y = (i//cols)*(th+header) + header + (th-im.height)//2
        canvas.paste(im,(x,y))
        d.text(((i%cols)*tw+5,(i//cols)*(th+header)+5), f'{i+1}. {p.name[-22:]}', fill='black', font=font)
    canvas.save(out/f'{folder}.jpg', quality=92)
