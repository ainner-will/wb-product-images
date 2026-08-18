from pathlib import Path
from PIL import Image, ImageOps, ImageDraw, ImageFont

base = Path(r"H:\New project\wb-product-images\images")
out = base / "work_preview"
folders = sorted([p for p in base.iterdir() if p.is_dir() and p.name not in {"work_preview", "outputs", "node_modules"}], key=lambda p:p.name)
thumb_w, thumb_h = 240, 300
font = ImageFont.load_default()
tiles = []
for folder in folders:
    imgs = sorted([p for p in folder.iterdir() if p.suffix.lower() in {".png",".jpg",".jpeg",".webp"}], key=lambda p:p.name)
    for idx, p in enumerate(imgs[:2]):
        try:
            im = Image.open(p).convert("RGB")
            im.thumbnail((thumb_w, thumb_h-35))
            canvas = Image.new("RGB", (thumb_w, thumb_h), "white")
            canvas.paste(im, ((thumb_w-im.width)//2, 25+(thumb_h-35-im.height)//2))
            d = ImageDraw.Draw(canvas)
            d.text((5,5), f"{folder.name} #{idx+1}", fill="black", font=font)
            tiles.append(canvas)
        except Exception:
            pass
cols=6
rows=(len(tiles)+cols-1)//cols
sheet=Image.new("RGB",(cols*thumb_w,rows*thumb_h),(230,230,230))
for i,t in enumerate(tiles):
    sheet.paste(t,((i%cols)*thumb_w,(i//cols)*thumb_h))
sheet.save(out/"sku_contact_sheet.jpg",quality=90)

# Detailed sheets: nine SKU folders per page, all images shown in rows.
detail_w, detail_h = 210, 285
for page in range(0, len(folders), 9):
    selected = folders[page:page+9]
    rows = []
    for folder in selected:
        imgs = sorted([p for p in folder.iterdir() if p.suffix.lower() in {".png",".jpg",".jpeg",".webp"}], key=lambda p:p.name)
        row_tiles=[]
        for idx,p in enumerate(imgs):
            im=Image.open(p).convert("RGB")
            im.thumbnail((detail_w,detail_h-30))
            c=Image.new("RGB",(detail_w,detail_h),"white")
            c.paste(im,((detail_w-im.width)//2,25+(detail_h-30-im.height)//2))
            ImageDraw.Draw(c).text((4,4),f"{folder.name} #{idx+1}",fill="black",font=font)
            row_tiles.append(c)
        rows.append(row_tiles)
    max_cols=max(len(x) for x in rows)
    d=Image.new("RGB",(max_cols*detail_w,len(rows)*detail_h),(225,225,225))
    for rr,row in enumerate(rows):
        for cc,t in enumerate(row): d.paste(t,(cc*detail_w,rr*detail_h))
    d.save(out/f"sku_details_{page//9+1}.jpg",quality=92)
