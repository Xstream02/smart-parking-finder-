$ErrorActionPreference = "Stop"

function Rgb($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

function Add-Rect($slide, $x, $y, $w, $h, $color, $transparency = 0, $radius = $false) {
    $shapeType = if ($radius) { 5 } else { 1 }
    $s = $slide.Shapes.AddShape($shapeType, $x, $y, $w, $h)
    $s.Fill.ForeColor.RGB = $color
    $s.Fill.Transparency = [single]$transparency
    $s.Line.Visible = 0
    return $s
}

function Add-Line($slide, $x1, $y1, $x2, $y2, $color, $weight = 2) {
    $s = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $s.Line.ForeColor.RGB = $color
    $s.Line.Weight = $weight
    return $s
}

function Add-Text($slide, $text, $x, $y, $w, $h, $size = 24, $color = $script:White, $bold = $false, $align = 1) {
    $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $tb.TextFrame.MarginLeft = 0
    $tb.TextFrame.MarginRight = 0
    $tb.TextFrame.MarginTop = 0
    $tb.TextFrame.MarginBottom = 0
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $text
    $tr.Font.Name = "Aptos Display"
    $tr.Font.Size = $size
    $tr.Font.Color.RGB = $color
    $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.ParagraphFormat.Alignment = $align
    return $tb
}

function Add-Bullet($slide, $text, $x, $y, $w, $accent, $size = 18) {
    Add-Rect $slide $x ($y + 8) 10 10 $accent 0 $true | Out-Null
    Add-Text $slide $text ($x + 22) $y $w 42 $size $script:Ink $false 1 | Out-Null
}

function Add-Card($slide, $x, $y, $w, $h, $title, $body, $accent) {
    Add-Rect $slide $x $y $w $h $script:Card 0 $true | Out-Null
    Add-Rect $slide $x $y 8 $h $accent 0 $false | Out-Null
    Add-Text $slide $title ($x + 22) ($y + 16) ($w - 40) 30 19 $script:Ink $true 1 | Out-Null
    Add-Text $slide $body ($x + 22) ($y + 52) ($w - 40) ($h - 62) 13 $script:Muted $false 1 | Out-Null
}

function Add-Footer($slide, $num) {
    Add-Text $slide "Smart Parking Finder" 44 508 260 18 10 $script:Muted $false 1 | Out-Null
    Add-Text $slide "$num / 7" 880 508 44 18 10 $script:Muted $false 3 | Out-Null
}

function New-Slide($presentation, $num, $title, $subtitle = "") {
    $slide = $presentation.Slides.Add($num, 12)
    Add-Rect $slide 0 0 960 540 $script:Bg 0 $false | Out-Null
    Add-Rect $slide 0 0 960 76 $script:TopBand 0 $false | Out-Null
    Add-Text $slide $title 44 24 620 40 28 $script:White $true 1 | Out-Null
    if ($subtitle -ne "") {
        Add-Text $slide $subtitle 690 32 226 24 12 $script:Soft $false 3 | Out-Null
    }
    Add-Footer $slide $num
    return $slide
}

$script:Bg = Rgb 9 17 31
$script:TopBand = Rgb 13 34 58
$script:White = Rgb 255 255 255
$script:Ink = Rgb 15 23 42
$script:Muted = Rgb 82 96 119
$script:Soft = Rgb 176 220 238
$script:Card = Rgb 248 251 255
$script:Cyan = Rgb 0 210 255
$script:Blue = Rgb 58 123 213
$script:Green = Rgb 34 197 94
$script:Amber = Rgb 245 158 11
$script:Pink = Rgb 244 63 94
$script:Purple = Rgb 139 92 246

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$outFile = Join-Path $root "SmartParkingFinder_Project_Presentation.pptx"
$imageFile = Join-Path $root "static\splash_car.png"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 960
$presentation.PageSetup.SlideHeight = 540

# Slide 1
$slide = $presentation.Slides.Add(1, 12)
Add-Rect $slide 0 0 960 540 $Bg 0 $false | Out-Null
Add-Rect $slide 0 0 960 540 (Rgb 4 19 35) 0 $false | Out-Null
if (Test-Path $imageFile) {
    $pic = $slide.Shapes.AddPicture($imageFile, 0, -1, 545, 0, 415, 540)
    $pic.Line.Visible = 0
}
Add-Rect $slide 44 66 490 348 (Rgb 255 255 255) 0.06 $true | Out-Null
Add-Text $slide "Smart Parking Finder" 72 98 450 74 40 $White $true 1 | Out-Null
Add-Text $slide "AI-assisted parking discovery, reservation, routing, and admin monitoring system" 74 188 405 64 20 $Soft $false 1 | Out-Null
Add-Rect $slide 74 285 166 42 $Cyan 0 $true | Out-Null
Add-Text $slide "Flask + SQLite" 96 296 126 20 14 $Ink $true 2 | Out-Null
Add-Rect $slide 256 285 184 42 $Green 0 $true | Out-Null
Add-Text $slide "Live WebSockets" 278 296 140 20 14 $Ink $true 2 | Out-Null
Add-Text $slide "Project Presentation" 74 452 240 20 14 $Soft $false 1 | Out-Null
Add-Text $slide "1 / 7" 872 508 44 18 10 $Soft $false 3 | Out-Null

# Slide 2
$slide = New-Slide $presentation 2 "Problem & Objective" "Why this project matters"
Add-Card $slide 54 118 260 142 "Parking Problem" "Drivers waste time finding free slots, slot status changes quickly, and manual checking is slow for both users and administrators." $Pink
Add-Card $slide 350 118 260 142 "Project Objective" "Provide a real-time parking finder where users can view available spots, reserve a slot, and navigate to the selected location." $Cyan
Add-Card $slide 646 118 260 142 "Admin Objective" "Give admins a dashboard to monitor slot status, approve paid bookings, verify OTPs, and track usage statistics." $Amber
Add-Text $slide "Core goal: reduce search time, avoid double booking, and make parking status visible instantly." 82 332 790 50 24 $White $true 2 | Out-Null
Add-Rect $slide 236 400 488 46 $Blue 0 $true | Out-Null
Add-Text $slide "Real-time updates + simple booking workflow + optional AI vision detection" 258 413 444 18 15 $White $true 2 | Out-Null

# Slide 3
$slide = New-Slide $presentation 3 "System Overview" "Main modules from your code"
Add-Card $slide 58 118 250 118 "User Interface" "templates/index.html, static/style.css`nSplash screen, login/register, vehicle selection, floor view, live map, and booking modal." $Cyan
Add-Card $slide 355 118 250 118 "Backend APIs" "app.py`nFlask routes for auth, parking data, status updates, relocation, and admin statistics." $Blue
Add-Card $slide 652 118 250 118 "Database" "database.db / SQLite`nUsers, ParkingSlots, and BookingsHistory tables store credentials, slot states, prices, OTP, and bookings." $Green
Add-Card $slide 58 294 250 118 "Realtime Layer" "Flask-SocketIO`nstatus_update and slots_relocated events keep user app and dashboard synchronized." $Purple
Add-Card $slide 355 294 250 118 "Admin Dashboard" "templates/admin.html`nSlot controls, booking approvals, rejection, release, utilization, and Chart.js monthly analysis." $Amber
Add-Card $slide 652 294 250 118 "Vision Module" "vision_module.py`nSimulated camera feed now, with YOLOv8 support path for real car detection using IoU." $Pink

# Slide 4
$slide = New-Slide $presentation 4 "User Workflow" "How a driver books a slot"
$steps = @(
    @("1", "Open App", "User starts from the splash screen and logs in or continues as guest."),
    @("2", "Select Vehicle", "Car or bike option leads the user into parking discovery."),
    @("3", "View Slots", "Leaflet map displays parking markers with live status and price."),
    @("4", "Book Slot", "User enters vehicle number, receives OTP, and confirms booking."),
    @("5", "Route", "App draws a route line from current location to the selected space.")
)
$x = 58
foreach ($s in $steps) {
    Add-Rect $slide $x 145 140 210 $Card 0 $true | Out-Null
    Add-Rect $slide ($x + 42) 118 56 56 $Cyan 0 $true | Out-Null
    Add-Text $slide $s[0] ($x + 60) 130 20 28 22 $Ink $true 2 | Out-Null
    Add-Text $slide $s[1] ($x + 16) 194 108 28 18 $Ink $true 2 | Out-Null
    Add-Text $slide $s[2] ($x + 16) 236 108 82 12 $Muted $false 2 | Out-Null
    if ($x -lt 730) { Add-Line $slide ($x + 142) 222 ($x + 174) 222 $Soft 2 | Out-Null }
    $x += 180
}
Add-Text $slide "Free slots are reserved instantly; paid slots move to Pending until admin approval." 120 410 720 30 20 $White $true 2 | Out-Null

# Slide 5
$slide = New-Slide $presentation 5 "Backend & Database Design" "Important implementation details"
Add-Text $slide "Flask API Endpoints" 66 116 300 28 22 $White $true 1 | Out-Null
Add-Bullet $slide "POST /api/auth/register and /api/auth/login for secure user access" 78 164 402 $Cyan 15
Add-Bullet $slide "GET /api/parking returns all slots with status, coordinates, price, vehicle, booking ID, and OTP" 78 212 402 $Green 15
Add-Bullet $slide "POST /api/parking/<id>/status changes Available, Occupied, Reserved, or Pending state" 78 270 402 $Amber 15
Add-Bullet $slide "POST /api/parking/relocate moves simulated slots around the user's live location" 78 328 402 $Purple 15
Add-Rect $slide 548 112 336 292 $Card 0 $true | Out-Null
Add-Text $slide "SQLite Tables" 574 136 260 24 22 $Ink $true 1 | Out-Null
Add-Text $slide "Users`n- UserID`n- Username`n- PasswordHash" 574 184 128 108 14 $Muted $false 1 | Out-Null
Add-Text $slide "ParkingSlots`n- SlotID, Location`n- Status, Lat, Lng`n- ReservedAt, Price`n- VehicleNo, BookingID, OTP" 714 184 148 152 14 $Muted $false 1 | Out-Null
Add-Text $slide "BookingsHistory`n- BookingID`n- Timestamp" 574 316 180 78 14 $Muted $false 1 | Out-Null
Add-Rect $slide 548 426 336 42 $Blue 0 $true | Out-Null
Add-Text $slide "Password hashing uses Werkzeug security helpers." 570 438 292 18 14 $White $true 2 | Out-Null

# Slide 6
$slide = New-Slide $presentation 6 "Realtime Admin & AI Vision" "Monitoring and automation"
Add-Card $slide 56 122 250 120 "Live Dashboard" "Admin page listens to Socket.IO events and refreshes slot rows, badges, counts, and utilization without page reloads." $Cyan
Add-Card $slide 356 122 250 120 "Approval Flow" "Paid bookings become Pending. Admin can approve, reject, mark occupied, or free an exited car." $Amber
Add-Card $slide 656 122 250 120 "Booking Security" "OTP and vehicle number are shown to admin, helping verify the driver before approving a paid slot." $Green
Add-Rect $slide 90 306 780 92 $Card 0 $true | Out-Null
Add-Text $slide "Vision Module Logic" 118 326 220 24 20 $Ink $true 1 | Out-Null
Add-Text $slide "The Python vision module simulates cars arriving/leaving every 10 seconds. It can also use YOLOv8 with a webcam, detect vehicles, calculate IoU against predefined parking boxes, then notify the Flask API." 118 362 700 36 14 $Muted $false 1 | Out-Null
Add-Rect $slide 160 430 640 42 $Purple 0 $true | Out-Null
Add-Text $slide "Camera/Simulation -> IoU Detection -> API Status Update -> Socket.IO Broadcast -> UI Refresh" 184 442 592 18 14 $White $true 2 | Out-Null

# Slide 7
$slide = New-Slide $presentation 7 "Features, Benefits & Future Scope" "Project conclusion"
Add-Text $slide "Key Features" 72 118 260 28 23 $White $true 1 | Out-Null
Add-Bullet $slide "Login/register and guest entry" 86 162 330 $Cyan 15
Add-Bullet $slide "Interactive Leaflet map with live markers" 86 202 330 $Green 15
Add-Bullet $slide "Free and paid parking slot support" 86 242 330 $Amber 15
Add-Bullet $slide "OTP-based booking verification" 86 282 330 $Purple 15
Add-Bullet $slide "Admin dashboard with Chart.js analysis" 86 322 330 $Pink 15
Add-Text $slide "Benefits" 520 118 260 28 23 $White $true 1 | Out-Null
Add-Bullet $slide "Reduces manual checking and waiting time" 534 162 330 $Cyan 15
Add-Bullet $slide "Avoids double booking with status locking" 534 202 330 $Green 15
Add-Bullet $slide "Makes slot availability transparent" 534 242 330 $Amber 15
Add-Bullet $slide "Supports practical campus/mall parking use cases" 534 282 330 $Purple 15
Add-Rect $slide 178 396 604 56 $Blue 0 $true | Out-Null
Add-Text $slide "Future scope: real payment gateway, QR check-in, improved route navigation, trained YOLO model, and multi-floor parking support." 204 412 552 20 15 $White $true 2 | Out-Null

$presentation.SaveAs($outFile)
$presentation.Close()
$ppt.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "Created $outFile"
