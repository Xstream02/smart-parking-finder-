$ErrorActionPreference = "Stop"

function XmlEscape($s) {
    return [System.Security.SecurityElement]::Escape([string]$s)
}

function Emu($pt) {
    return [int64]([double]$pt * 12700)
}

function HexColor($hex) {
    return $hex.TrimStart("#").ToUpperInvariant()
}

function TextRuns($lines, $fontSize, $color, $bold) {
    $b = if ($bold) { ' b="1"' } else { "" }
    $sz = [int]($fontSize * 100)
    $safeColor = HexColor $color
    $paras = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        $paras.Add("<a:p><a:r><a:rPr lang=""en-US"" sz=""$sz""$b><a:solidFill><a:srgbClr val=""$safeColor""/></a:solidFill><a:latin typeface=""Aptos""/></a:rPr><a:t>$(XmlEscape $line)</a:t></a:r><a:endParaRPr lang=""en-US"" sz=""$sz""/></a:p>")
    }
    return ($paras -join "")
}

function ShapeXml($id, $name, $x, $y, $w, $h, $fill, $textLines, $fontSize, $textColor, $bold = $false, $radius = $true) {
    $geom = if ($radius) { "roundRect" } else { "rect" }
    $body = if ($textLines) { TextRuns $textLines $fontSize $textColor $bold } else { "<a:p/>" }
    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$id" name="$(XmlEscape $name)"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="$(Emu $x)" y="$(Emu $y)"/><a:ext cx="$(Emu $w)" cy="$(Emu $h)"/></a:xfrm>
    <a:prstGeom prst="$geom"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="$(HexColor $fill)"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91440" tIns="60960" rIns="91440" bIns="60960" anchor="mid"/>
    <a:lstStyle/>
    $body
  </p:txBody>
</p:sp>
"@
}

function TextBoxXml($id, $name, $x, $y, $w, $h, $textLines, $fontSize, $textColor, $bold = $false) {
    $body = TextRuns $textLines $fontSize $textColor $bold
    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$id" name="$(XmlEscape $name)"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="$(Emu $x)" y="$(Emu $y)"/><a:ext cx="$(Emu $w)" cy="$(Emu $h)"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/><a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/>
    <a:lstStyle/>
    $body
  </p:txBody>
</p:sp>
"@
}

function SlideXml($title, $subtitle, $elements) {
    $all = New-Object System.Collections.Generic.List[string]
    $all.Add((ShapeXml 2 "Background" 0 0 960 540 "#09111F" $null 1 "#FFFFFF" $false $false))
    $all.Add((ShapeXml 3 "Header" 0 0 960 76 "#0D223A" $null 1 "#FFFFFF" $false $false))
    $all.Add((TextBoxXml 4 "Title" 44 22 660 42 @($title) 28 "#FFFFFF" $true))
    if ($subtitle) { $all.Add((TextBoxXml 5 "Subtitle" 690 30 220 26 @($subtitle) 12 "#B0DCEE" $false)) }
    foreach ($e in $elements) { $all.Add($e) }
    $content = $all -join "`n"
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      $content
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function WriteUtf8($path, $content) {
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $content, $enc)
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp = Join-Path $root "_pptx_compat_tmp"
$out = Join-Path $root "SmartParkingFinder_Project_Presentation_COMPAT.pptx"
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
if (Test-Path $out) { Remove-Item $out -Force }

$dirs = @(
    "_rels","docProps","ppt","ppt\_rels","ppt\slides","ppt\slides\_rels",
    "ppt\slideMasters","ppt\slideMasters\_rels","ppt\slideLayouts","ppt\theme"
)
foreach ($d in $dirs) { New-Item -ItemType Directory -Path (Join-Path $tmp $d) | Out-Null }

$slides = @()

$e = @()
$e += ShapeXml 10 "HeroCard" 58 112 512 286 "#F8FBFF" @("AI-assisted parking discovery", "Real-time slot status, booking, routing, and admin monitoring in one project.") 22 "#0F172A" $true $true
$e += ShapeXml 11 "Tag1" 88 318 138 42 "#00D2FF" @("Flask") 16 "#0F172A" $true $true
$e += ShapeXml 12 "Tag2" 246 318 138 42 "#22C55E" @("SQLite") 16 "#0F172A" $true $true
$e += ShapeXml 13 "Tag3" 404 318 138 42 "#F59E0B" @("Socket.IO") 16 "#0F172A" $true $true
$e += ShapeXml 14 "Side" 628 116 250 278 "#3A7BD5" @("Smart Parking Finder", "Project Presentation", "Built from your codebase") 24 "#FFFFFF" $true $true
$slides += SlideXml "Smart Parking Finder" "Project Presentation" $e

$e = @()
$e += ShapeXml 10 "Problem" 58 126 250 142 "#F8FBFF" @("Problem", "Drivers waste time finding slots and parking status changes quickly.") 18 "#0F172A" $true $true
$e += ShapeXml 11 "Objective" 355 126 250 142 "#F8FBFF" @("Objective", "Show available spaces, support reservation, and guide users to the selected slot.") 18 "#0F172A" $true $true
$e += ShapeXml 12 "Admin" 652 126 250 142 "#F8FBFF" @("Admin Need", "Monitor slots, approve paid bookings, verify OTP, and view platform usage.") 18 "#0F172A" $true $true
$e += ShapeXml 13 "Goal" 176 344 608 72 "#00D2FF" @("Goal: reduce search time, avoid double booking, and make availability visible instantly.") 20 "#0F172A" $true $true
$slides += SlideXml "Problem & Objective" "Why this project matters" $e

$e = @()
$modules = @(
    @("User Interface","index.html and style.css: splash, login, vehicle choice, map, booking modal.","#00D2FF"),
    @("Backend APIs","app.py: auth, parking data, status updates, relocation, admin stats.","#3A7BD5"),
    @("Database","SQLite tables: Users, ParkingSlots, BookingsHistory.","#22C55E"),
    @("Realtime Layer","Flask-SocketIO broadcasts status_update and slots_relocated.","#8B5CF6"),
    @("Admin Dashboard","admin.html: controls, approvals, utilization, Chart.js report.","#F59E0B"),
    @("Vision Module","vision_module.py: simulated feed plus YOLOv8-ready detection path.","#F43F5E")
)
$i = 0
foreach ($m in $modules) {
    $x = 58 + (($i % 3) * 297)
    $y = 116 + ([math]::Floor($i / 3) * 178)
    $e += ShapeXml (10+$i) $m[0] $x $y 250 128 $m[2] @($m[0], $m[1]) 15 "#FFFFFF" $true $true
    $i++
}
$slides += SlideXml "System Overview" "Main modules from your code" $e

$e = @()
$steps = @("Open App","Login or guest entry","View live map","Enter vehicle number and OTP","Reserve and route")
for ($i=0; $i -lt $steps.Count; $i++) {
    $x = 58 + ($i * 180)
    $e += ShapeXml (10+$i) "Step$i" $x 158 140 178 "#F8FBFF" @(("$($i+1). " + $steps[$i])) 17 "#0F172A" $true $true
}
$e += ShapeXml 20 "Note" 140 408 680 48 "#22C55E" @("Free slots reserve instantly; paid slots wait for admin approval.") 18 "#0F172A" $true $true
$slides += SlideXml "User Workflow" "How a driver books a slot" $e

$e = @()
$e += ShapeXml 10 "Apis" 54 116 400 322 "#F8FBFF" @("Flask API Endpoints", "POST /api/auth/register and /api/auth/login", "GET /api/parking", "POST /api/parking/<id>/status", "POST /api/parking/relocate", "GET /api/admin/stats") 15 "#0F172A" $true $true
$e += ShapeXml 11 "Db" 508 116 398 322 "#F8FBFF" @("SQLite Database", "Users: UserID, Username, PasswordHash", "ParkingSlots: SlotID, Location, Status, Lat, Lng, Price, VehicleNo, BookingID, OTP", "BookingsHistory: BookingID, Timestamp") 15 "#0F172A" $true $true
$e += ShapeXml 12 "Security" 236 462 488 38 "#8B5CF6" @("Passwords are stored with Werkzeug hashing helpers.") 15 "#FFFFFF" $true $true
$slides += SlideXml "Backend & Database Design" "Important implementation details" $e

$e = @()
$e += ShapeXml 10 "Dashboard" 58 124 250 128 "#00D2FF" @("Live Dashboard", "Slot cards, availability counts, utilization, and booking chart update from API data.") 15 "#0F172A" $true $true
$e += ShapeXml 11 "Approval" 355 124 250 128 "#F59E0B" @("Approval Flow", "Paid slots become Pending. Admin approves, rejects, occupies, or frees slots.") 15 "#0F172A" $true $true
$e += ShapeXml 12 "Otp" 652 124 250 128 "#22C55E" @("OTP Verification", "Booking ID, vehicle number, and OTP help confirm the correct user.") 15 "#0F172A" $true $true
$e += ShapeXml 13 "Vision" 90 322 780 88 "#F8FBFF" @("Vision Module", "Simulation checks six predefined parking boxes every 10 seconds. YOLOv8 can be enabled for webcam vehicle detection, then IoU decides whether a slot is Occupied or Available.") 15 "#0F172A" $true $true
$slides += SlideXml "Realtime Admin & AI Vision" "Monitoring and automation" $e

$e = @()
$e += ShapeXml 10 "Features" 58 116 390 294 "#F8FBFF" @("Key Features", "Login/register and guest mode", "Interactive Leaflet map", "Free and paid parking support", "OTP booking verification", "Admin dashboard with analytics") 15 "#0F172A" $true $true
$e += ShapeXml 11 "Benefits" 512 116 390 294 "#F8FBFF" @("Benefits", "Saves parking search time", "Reduces manual checking", "Avoids double booking", "Improves visibility for admins", "Useful for campus, mall, and office parking") 15 "#0F172A" $true $true
$e += ShapeXml 12 "Future" 156 448 648 42 "#3A7BD5" @("Future scope: payment gateway, QR check-in, trained YOLO model, multi-floor support.") 15 "#FFFFFF" $true $true
$slides += SlideXml "Features, Benefits & Future Scope" "Project conclusion" $e

for ($i=0; $i -lt $slides.Count; $i++) {
    WriteUtf8 (Join-Path $tmp "ppt\slides\slide$($i+1).xml") $slides[$i]
    WriteUtf8 (Join-Path $tmp "ppt\slides\_rels\slide$($i+1).xml.rels") "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"" Target=""../slideLayouts/slideLayout1.xml""/></Relationships>"
}

$slideOverrides = (1..7 | ForEach-Object { "<Override PartName=""/ppt/slides/slide$_.xml"" ContentType=""application/vnd.openxmlformats-officedocument.presentationml.slide+xml""/>" }) -join "`n"
WriteUtf8 (Join-Path $tmp "[Content_Types].xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  $slideOverrides
</Types>
"@

WriteUtf8 (Join-Path $tmp "_rels\.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$slideIds = (1..7 | ForEach-Object { "<p:sldId id=""$([int](255 + $_))"" r:id=""rId$_""/>" }) -join "`n"
WriteUtf8 (Join-Path $tmp "ppt\presentation.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId8"/></p:sldMasterIdLst>
  <p:sldIdLst>$slideIds</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@

$rels = (1..7 | ForEach-Object { "<Relationship Id=""rId$_"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide$_.xml""/>" }) -join "`n"
WriteUtf8 (Join-Path $tmp "ppt\_rels\presentation.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $rels
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
"@

WriteUtf8 (Join-Path $tmp "ppt\slideMasters\slideMaster1.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>
"@
WriteUtf8 (Join-Path $tmp "ppt\slideMasters\_rels\slideMaster1.xml.rels") "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"" Target=""../slideLayouts/slideLayout1.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme1.xml""/></Relationships>"

WriteUtf8 (Join-Path $tmp "ppt\slideLayouts\slideLayout1.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"@
New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\slideLayouts\_rels") | Out-Null
WriteUtf8 (Join-Path $tmp "ppt\slideLayouts\_rels\slideLayout1.xml.rels") "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"" Target=""../slideMasters/slideMaster1.xml""/></Relationships>"

WriteUtf8 (Join-Path $tmp "ppt\theme\theme1.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements><a:clrScheme name="Custom"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F8FBFF"/></a:lt2><a:accent1><a:srgbClr val="00D2FF"/></a:accent1><a:accent2><a:srgbClr val="3A7BD5"/></a:accent2><a:accent3><a:srgbClr val="22C55E"/></a:accent3><a:accent4><a:srgbClr val="F59E0B"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="F43F5E"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>
</a:theme>
"@

WriteUtf8 (Join-Path $tmp "docProps\app.xml") "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes""><Application>Codex</Application><PresentationFormat>On-screen Show</PresentationFormat><Slides>7</Slides></Properties>"
WriteUtf8 (Join-Path $tmp "docProps\core.xml") "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""><dc:title>Smart Parking Finder Project Presentation</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy></cp:coreProperties>"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($out, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    Get-ChildItem $tmp -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($tmp.Length + 1).Replace("\", "/")
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
    }
}
finally {
    $zip.Dispose()
}
Remove-Item $tmp -Recurse -Force
Write-Output "Created $out"
