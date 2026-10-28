$ErrorActionPreference = 'Stop'
$source = Get-ChildItem -File -Filter '*sku_list.csv' | Select-Object -First 1
$rows = @(Import-Csv -LiteralPath $source.FullName -Encoding 936)
$colors = [ordered]@{'黑色'=@('BLK','BLACK');'白色'=@('WHT','WHITE');'宝蓝色'=@('RBL','ROYAL BLUE');'深棕色'=@('DBR','DARK BROWN');'粉红色'=@('PNK','PINK');'玫红色'=@('FUS','FUCHSIA');'香槟色'=@('CHP','CHAMPAGNE');'紫色'=@('PUR','PURPLE')}
foreach($dir in @('SKU','SKU_EN')) { if(Test-Path -LiteralPath $dir){throw "目标目录已存在：$dir"} }
$items = foreach($r in $rows){
 $cn = @($colors.Keys | Where-Object {$r.'颜色'.StartsWith($_)})
 if($cn.Count -ne 1){throw '颜色匹配失败'}
 $c = $colors[$cn[0]]
 $size = ($r.'尺寸' -replace '[×xX_*]','X') -replace '[^0-9X-]',''
 $bag = if($r.'颜色'.Contains('带包')){'BAG'}else{'NB'}
 $code = "MT-$size-$($c[0])-HA-$bag"
 $english = "ADJUSTABLE FOLDING MASSAGE TABLE, $($c[1]), $size CM, HEADREST AND ARMRESTS"
 if($bag -eq 'BAG'){$english += ', WITH CARRY BAG'}else{$english += ', CARRY BAG NOT SPECIFIED'}
 [pscustomobject]@{'SKU ID'=$r.'SKU ID';'中文SKU名称'=$r.'SKU名称';'英文名称'=$english;'英文商品编码'=$code;'颜色中文'=$cn[0];'颜色英文'=$c[1];'尺寸'=$size;'包配置'=$(if($bag -eq 'BAG'){'带包'}else{'原名称未注明带包'});'图片URL'=$r.'SKU 图片';'原图路径'="SKU/$($r.'SKU ID').jpg";'英文文件夹图片路径'="SKU_EN/$code/$code.jpg"}
}
if(@($items.'英文商品编码' | Sort-Object -Unique).Count -ne $rows.Count){throw '编码重复'}
if(@($items | Where-Object {$_.'英文商品编码' -cnotmatch '^[A-Z0-9-]+$'}).Count){throw '编码格式无效'}
New-Item -ItemType Directory -Path 'SKU','SKU_EN' | Out-Null
$cache = @{}
foreach($item in $items){
 $url=$item.'图片URL'; $dest=Join-Path $PWD $item.'原图路径'
 if($cache.ContainsKey($url)){Copy-Item -LiteralPath $cache[$url] -Destination $dest}else{
  $ok=$false
  for($attempt=1;$attempt -le 3;$attempt++){
   try {Invoke-WebRequest -Uri $url -OutFile $dest -TimeoutSec 45; $bytes=[IO.File]::ReadAllBytes($dest); if($bytes.Length -lt 100 -or $bytes[0] -ne 255 -or $bytes[1] -ne 216){throw '不是有效JPEG图片'}; $ok=$true;break}catch{if($attempt -eq 3){throw}}
  }
  $cache[$url]=$dest
 }
 $folder=Join-Path 'SKU_EN' $item.'英文商品编码'
 New-Item -ItemType Directory -Path $folder | Out-Null
 Copy-Item -LiteralPath $dest -Destination (Join-Path $PWD $item.'英文文件夹图片路径')
}
$items | Export-Csv -LiteralPath 'SKU中英文对照表.csv' -NoTypeInformation -Encoding utf8BOM
$intro = '<meta charset="utf-8"><style>body{font:14px Arial;padding:24px}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:8px;text-align:left}th{background:#eaf1fa}img{width:120px}</style><h1>按摩床 SKU 中英文对照表</h1><p>MT=按摩床；HA=头枕和扶手；BAG=带包；NB=原名称未注明带包。尺寸单位：CM；50-72 为高度范围。颜色编码：BLK 黑色、WHT 白色、RBL 宝蓝色、DBR 深棕色、PNK 粉红色、FUS 玫红色、CHP 香槟色、PUR 紫色。</p>'
$html = $items | Select-Object 'SKU ID','中文SKU名称','英文名称','英文商品编码','图片URL','英文文件夹图片路径' | ConvertTo-Html -Fragment
Set-Content -LiteralPath 'SKU中英文对照表.html' -Value ($intro + ($html -join "`n")) -Encoding utf8BOM
$checks = foreach($item in $items){
 $a=Get-Item -LiteralPath $item.'原图路径';$b=Get-Item -LiteralPath $item.'英文文件夹图片路径'
 if((Get-FileHash -LiteralPath $a.FullName).Hash -ne (Get-FileHash -LiteralPath $b.FullName).Hash){throw '图片复制校验失败'}
}
[pscustomobject]@{SKU_COUNT=$items.Count;UNIQUE_URLS=$cache.Count;ORIGINAL_IMAGES=@(Get-ChildItem SKU -File).Count;ENGLISH_FOLDERS=@(Get-ChildItem SKU_EN -Directory).Count;ENGLISH_IMAGES=@(Get-ChildItem SKU_EN -Recurse -File).Count;CODES_UNIQUE=$true} | ConvertTo-Json
