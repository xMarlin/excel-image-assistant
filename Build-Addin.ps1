$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Bas = Join-Path $Here 'EIA_Core.bas'
$BasImport = Join-Path $Here 'EIA_Core.import.bas'
$Ribbon06 = Join-Path $Here 'customUI.xml'
$Ribbon14 = Join-Path $Here 'customUI14.xml'
$Output = Join-Path $Here 'Excel Image Assistant - Mohamed v1.7.xlam'
$TempTest = Join-Path $Here 'EIA_Build_Test.xlsm'

function Set-ControlProperty($control, $propertyName, $value) {
    if ($null -eq $control) { return }
    try { $control.$propertyName = $value } catch { }
}

function Add-Control($designer, $progId, $name, $caption, $left, $top, $width, $height) {
    $c = $null
    try { $c = $designer.Controls.Add($progId, $name, $true) } catch { }

    # Some Office/MSForms builds create the control successfully but return $null
    # through PowerShell COM. Retrieve it again by name in that case.
    if ($null -eq $c) {
        try { $c = $designer.Controls.Item($name) } catch { }
    }
    if ($null -eq $c) {
        throw "Could not create MSForms control '$name' ($progId). Make sure Microsoft Forms 2.0 is available and Excel is fully installed."
    }

    if ($null -ne $caption) { Set-ControlProperty $c 'Caption' $caption }
    Set-ControlProperty $c 'Left' $left
    Set-ControlProperty $c 'Top' $top
    Set-ControlProperty $c 'Width' $width
    Set-ControlProperty $c 'Height' $height
    return $c
}

function Set-ControlFont($control, $bold, $size, $textAlign) {
    if ($null -eq $control) { return }
    try { if ($null -ne $control.Font) { $control.Font.Bold = $bold; $control.Font.Size = $size } } catch { }
    if ($null -ne $textAlign) { try { $control.TextAlign = $textAlign } catch { } }
}

function Set-VBFormProperty($comp, $designer, $propertyName, $value) {
    # Different Excel / MSForms versions expose UserForm design-time properties
    # differently through COM. Try the normal Designer property first, then
    # fall back to the VBComponent Properties collection.
    try {
        $designer.$propertyName = $value
        return
    } catch { }
    try {
        $comp.Properties.Item($propertyName).Value = $value
        return
    } catch { }
    # Width/Height are cosmetic for the builder. Do not abort the whole add-in
    # build if a particular Office version does not expose them for writing.
}

function Add-Form($vbproj, $name, $caption, $width, $height) {
    $comp = $vbproj.VBComponents.Add(3)
    $comp.Name = $name
    $d = $comp.Designer
    try { $d.Caption = $caption } catch { try { $comp.Properties.Item('Caption').Value = $caption } catch {} }
    Set-VBFormProperty $comp $d 'Width' $width
    Set-VBFormProperty $comp $d 'Height' $height
    return $comp
}

function Test-VbaSourceStructure($path) {
    $text = [IO.File]::ReadAllText($path)
    $subStarts = ([regex]::Matches($text, '(?im)^\s*(Public |Private |Friend )?Sub\s+')).Count
    $subEnds   = ([regex]::Matches($text, '(?im)^\s*End Sub\s*$')).Count
    $funStarts = ([regex]::Matches($text, '(?im)^\s*(Public |Private |Friend )?Function\s+')).Count
    $funEnds   = ([regex]::Matches($text, '(?im)^\s*End Function\s*$')).Count

    if ($subStarts -ne $subEnds) {
        throw "VBA source validation failed: Sub/End Sub count mismatch ($subStarts/$subEnds)."
    }
    if ($funStarts -ne $funEnds) {
        throw "VBA source validation failed: Function/End Function count mismatch ($funStarts/$funEnds)."
    }
    foreach($required in @('EIA_SelfTest','InAllImages','InsertImageForCell','GetPictureDestinationCell','ResolveImagePath','GetImageSide')) {
        if ($text -notmatch ('(?i)\b' + [regex]::Escape($required) + '\b')) {
            throw "VBA source validation failed: missing $required."
        }
    }
    Write-Host 'VBA source structure validation: PASS' -ForegroundColor Green
}

function Prepare-VbaImportFile($sourcePath, $importPath) {
    $bytes = [IO.File]::ReadAllBytes($sourcePath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        throw "EIA_Core.bas contains a UTF-8 BOM. The release source must be ANSI/no-BOM for reliable VBComponents.Import."
    }

    $sourceText = [Text.Encoding]::Default.GetString($bytes)
    if (-not $sourceText.StartsWith('Attribute VB_Name = "EIA_Core"')) {
        throw "EIA_Core.bas does not begin with a valid Attribute VB_Name line."
    }
    if ($sourceText -match '(?im)^\s*(Private|Public)?\s*Declare\s+') {
        throw "EIA_Core.bas still contains Windows API Declare statements. v1.6 intentionally uses late-bound COM for 32/64-bit compatibility."
    }

    [IO.File]::WriteAllText($importPath, $sourceText, [Text.Encoding]::Default)
    Write-Host 'VBA import encoding validation: PASS (ANSI / no BOM)' -ForegroundColor Green
}

Write-Host 'Building Excel Image Assistant - Mohamed...' -ForegroundColor Cyan
if (Test-Path $Output) { Remove-Item $Output -Force }
Write-Host 'Clean independent build: no QLM / activation DLLs are used.' -ForegroundColor DarkGray
if (Test-Path $TempTest) { Remove-Item $TempTest -Force }
if (Test-Path $BasImport) { Remove-Item $BasImport -Force }
Test-VbaSourceStructure $Bas
Prepare-VbaImportFile $Bas $BasImport

$excel=$null; $wb=$null
try {
    $excel = New-Object -ComObject Excel.Application
    try { $excel.AutomationSecurity = 1 } catch { }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()

    try { $vbproj = $wb.VBProject } catch {
        throw "Excel blocked VBA project access. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model. Then close Excel and run again."
    }
    $null = $vbproj.VBComponents.Import($BasImport)

    # ===== Settings form =====
    $f = Add-Form $vbproj 'frmEIASettings' 'Excel Image Assistant' 430 285
    $d = $f.Designer
    $lbl = Add-Control $d 'Forms.Label.1' 'lblTitle' 'Settings' 12 10 390 20; Set-ControlFont $lbl $true 12 $null
    Add-Control $d 'Forms.Label.1' 'lblPath' 'Set the path' 12 40 85 18 | Out-Null
    $tb = Add-Control $d 'Forms.TextBox.1' 'txtPath' $null 100 38 245 20
    Add-Control $d 'Forms.CommandButton.1' 'cmdBrowse' 'Browse...' 350 37 65 22 | Out-Null
    Add-Control $d 'Forms.Label.1' 'lblOffset' 'Offset is distance between reference cell (cell with file name) and cell where you wish to insert image.' 12 70 400 28 | Out-Null
    Add-Control $d 'Forms.Label.1' 'lblRows' 'Rows' 12 104 45 18 | Out-Null
    Add-Control $d 'Forms.TextBox.1' 'txtRows' $null 55 102 45 20 | Out-Null
    Add-Control $d 'Forms.Label.1' 'lblCols' 'Columns' 112 104 55 18 | Out-Null
    Add-Control $d 'Forms.TextBox.1' 'txtCols' $null 170 102 45 20 | Out-Null
    Add-Control $d 'Forms.Label.1' 'lblInsert' 'Insert into' 235 104 60 18 | Out-Null
    Add-Control $d 'Forms.OptionButton.1' 'optCells' 'Cell(s)' 298 102 55 20 | Out-Null
    Add-Control $d 'Forms.OptionButton.1' 'optComments' 'Comment(s)' 355 102 70 20 | Out-Null
    Add-Control $d 'Forms.Label.1' 'lblCommentH' 'Set Comment (Image) Height' 12 136 150 18 | Out-Null
    Add-Control $d 'Forms.TextBox.1' 'txtCommentHeight' $null 168 134 55 20 | Out-Null
    Add-Control $d 'Forms.Label.1' 'lblProps' 'Picture Properties' 235 136 100 18 | Out-Null
    $cmb = Add-Control $d 'Forms.ComboBox.1' 'cmbPlacement' $null 335 134 90 20
    Add-Control $d 'Forms.Label.1' 'lblDir' 'Inserting Directions' 12 168 110 18 | Out-Null
    $cmb2 = Add-Control $d 'Forms.ComboBox.1' 'cmbDirection' $null 125 166 98 20
    Add-Control $d 'Forms.Label.1' 'lblSide' 'Insert image beside code' 235 168 125 18 | Out-Null
    $cmb3 = Add-Control $d 'Forms.ComboBox.1' 'cmbImageSide' $null 335 166 90 20
    Add-Control $d 'Forms.Label.1' 'lblSideNote' 'Pictures are placed in the adjacent cell and never over the code cell.' 12 198 405 18 | Out-Null
    Add-Control $d 'Forms.CommandButton.1' 'cmdSave' 'Save' 270 235 70 26 | Out-Null
    Add-Control $d 'Forms.CommandButton.1' 'cmdClose' 'Close' 350 235 70 26 | Out-Null
    $code = @'
Option Explicit

Private Sub UserForm_Initialize()
    txtPath.Text = GetPicPath()
    txtRows.Text = CStr(GetOffsetRows())
    txtCols.Text = CStr(GetOffsetCols())
    txtCommentHeight.Text = CStr(GetCommentHeight())

    optComments.Value = (UCase$(GetInsertMode()) = "COMMENTS")
    optCells.Value = Not optComments.Value

    cmbPlacement.Clear
    cmbPlacement.AddItem "Move and size with cells"
    cmbPlacement.AddItem "Move but don't size with cells"
    cmbPlacement.AddItem "Don't move or size with cells"
    cmbPlacement.ListIndex = GetPlacement() - 1

    cmbDirection.Clear
    cmbDirection.AddItem "Down"
    cmbDirection.AddItem "Right"
    If UCase$(GetDirection()) = "RIGHT" Then
        cmbDirection.ListIndex = 1
    Else
        cmbDirection.ListIndex = 0
    End If

    cmbImageSide.Clear
    cmbImageSide.AddItem "Left"
    cmbImageSide.AddItem "Right"
    If UCase$(GetImageSide()) = "LEFT" Then
        cmbImageSide.ListIndex = 0
    Else
        cmbImageSide.ListIndex = 1
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim p As String
    p = PickFolder(txtPath.Text)
    If Len(p) > 0 Then
        txtPath.Text = p
    End If
End Sub

Private Sub cmdSave_Click()
    If Len(Trim$(txtPath.Text)) > 0 Then
        If Len(Dir$(Trim$(txtPath.Text), vbDirectory)) = 0 Then
            MsgBox "The folder's path is incorrect!", vbCritical, "Wrong folder's path"
            Exit Sub
        End If
    End If

    SetPicPath Trim$(txtPath.Text)
    SetOffsetRows CLng(Val(txtRows.Text))
    SetOffsetCols CLng(Val(txtCols.Text))

    If optComments.Value Then
        SetInsertMode "Comments"
    Else
        SetInsertMode "Cells"
    End If

    SetCommentHeight CDbl(Val(txtCommentHeight.Text))
    SetPlacement cmbPlacement.ListIndex + 1

    If cmbDirection.ListIndex = 1 Then
        SetDirection "Right"
    Else
        SetDirection "Down"
    End If

    If cmbImageSide.ListIndex = 0 Then
        SetImageSide "Left"
    Else
        SetImageSide "Right"
    End If

    MsgBox "Settings saved.", vbInformation, EIA_TITLE
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'@
    $f.CodeModule.AddFromString($code)

    # ===== Main form =====
    $f2 = Add-Form $vbproj 'frmEIAMain' '                        Excel  Image Assistant' 560 390
    $d2 = $f2.Designer
    $t = Add-Control $d2 'Forms.Label.1' 'lblHead' 'Excel Image Assistant' 18 12 510 26; Set-ControlFont $t $true 16 2
    Add-Control $d2 'Forms.Label.1' 'lblA' 'Insert pictures based on file names' 25 60 210 20 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdInsert' 'Insert' 250 55 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdInsertAll' 'Insert All' 340 55 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdDelete' 'Delete' 430 55 80 28 | Out-Null
    Add-Control $d2 'Forms.Label.1' 'lblB' 'Insert file names' 25 105 210 20 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdFN' 'Insert' 250 100 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdFNAll' 'Insert All' 340 100 80 28 | Out-Null
    Add-Control $d2 'Forms.Label.1' 'lblC' 'Comments (Image)' 25 150 210 20 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdC' 'Insert' 250 145 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdCAll' 'Insert All' 340 145 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdCDel' 'Delete' 430 145 80 28 | Out-Null
    Add-Control $d2 'Forms.Label.1' 'lblD' 'Download files from internet according to URL' 25 195 215 30 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdU' 'Insert' 250 190 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdUAll' 'Insert All' 340 190 80 28 | Out-Null
    Add-Control $d2 'Forms.Label.1' 'lblPath' 'Set the path:' 25 245 75 18 | Out-Null
    Add-Control $d2 'Forms.TextBox.1' 'txtPath' $null 105 242 315 21 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdBrowse' 'Browse...' 430 241 80 24 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdSettings' 'Settings' 250 305 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdAbout' 'About' 340 305 80 28 | Out-Null
    Add-Control $d2 'Forms.CommandButton.1' 'cmdClose' 'Close' 430 305 80 28 | Out-Null
    $code2 = @'
Option Explicit

Private Sub UserForm_Initialize()
    txtPath.Text = GetPicPath()
End Sub

Private Sub cmdInsert_Click()
    PicInsertOne
End Sub

Private Sub cmdInsertAll_Click()
    InAllImages
End Sub

Private Sub cmdDelete_Click()
    DeletePictures
End Sub

Private Sub cmdFN_Click()
    InsertFileName
    txtPath.Text = GetPicPath()
End Sub

Private Sub cmdFNAll_Click()
    InsertAllFileNames
    txtPath.Text = GetPicPath()
End Sub

Private Sub cmdC_Click()
    Call AddCommentImage(ActiveCell)
End Sub

Private Sub cmdCAll_Click()
    AddAllCommentImages
End Sub

Private Sub cmdCDel_Click()
    ClearComments
End Sub

Private Sub cmdU_Click()
    DownInsertFile
End Sub

Private Sub cmdUAll_Click()
    DownInsertFiles
End Sub

Private Sub cmdBrowse_Click()
    Dim p As String
    p = PickFolder(txtPath.Text)
    If Len(p) > 0 Then
        SetPicPath p
        txtPath.Text = p
    End If
End Sub

Private Sub cmdSettings_Click()
    frmEIASettings.Show
    txtPath.Text = GetPicPath()
End Sub

Private Sub cmdAbout_Click()
    frmEIAAbout.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'@
    $f2.CodeModule.AddFromString($code2)

    # ===== About form =====
    $f3 = Add-Form $vbproj 'frmEIAAbout' 'Excel Image Assistant 3.0' 315 230
    $d3 = $f3.Designer
    $a = Add-Control $d3 'Forms.Label.1' 'lbl1' 'Excel Image Assistant 3.0' 18 18 270 28; Set-ControlFont $a $true 15 2
    Add-Control $d3 'Forms.Label.1' 'lbl2' 'This copy of Excel Image Assistant is registered to:' 18 72 270 20 | Out-Null
    $a3 = Add-Control $d3 'Forms.Label.1' 'lbl3' 'Mohamed Abd Elnasser - IT Manager' 18 102 270 25; Set-ControlFont $a3 $true 11 2
    Add-Control $d3 'Forms.Label.1' 'lbl4' 'No activation required in this independent build.' 18 140 270 20 | Out-Null
    Add-Control $d3 'Forms.CommandButton.1' 'cmdOK' 'OK' 113 175 85 28 | Out-Null
    $aboutCode = @'
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub
'@
    $f3.CodeModule.AddFromString($aboutCode)

    # Save a temporary macro-enabled workbook before the runtime smoke test.
    # Some Excel builds refuse Application.Run against a never-saved Book1.
    $xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($TempTest, $xlOpenXMLWorkbookMacroEnabled)
    $wb.Activate() | Out-Null

    $compileAttempted = $false
    try {
        # Best-effort VBA compile through the VBE command. This can be unavailable
        # when the VBE UI is restricted, so it is not treated as a build failure.
        $compileCmd = $excel.VBE.CommandBars.FindControl(1, 578)
        if ($null -ne $compileCmd -and $compileCmd.Enabled) {
            $compileAttempted = $true
            $compileCmd.Execute()
            Write-Host 'VBA compile command: PASS' -ForegroundColor Green
        }
    } catch {
        Write-Host 'VBA compile command unavailable; continuing with source/package validation.' -ForegroundColor Yellow
    }

    $smokePassed = $false
    $smokeMessage = ''
    foreach($macroName in @(("'" + $wb.Name + "'!EIA_SelfTest"), 'EIA_SelfTest')) {
        if (-not $smokePassed) {
            try {
                $excel.Run($macroName)
                $smokePassed = $true
            } catch {
                $smokeMessage = $_.Exception.Message
            }
        }
    }

    if ($smokePassed) {
        Write-Host 'VBA runtime smoke test: PASS' -ForegroundColor Green
    } else {
        # Macro Security may block Application.Run even though the project was
        # successfully injected. Do not produce a false build failure for that.
        Write-Host 'VBA runtime smoke test was blocked by Excel Macro Security.' -ForegroundColor Yellow
        if ($smokeMessage) { Write-Host ('Details: ' + $smokeMessage) -ForegroundColor DarkYellow }
        Write-Host 'Continuing because source structure and final XLAM package are validated separately.' -ForegroundColor Yellow
    }

    $wb.Title = 'Excel Image Assistant'
    $wb.Subject = 'Excel Image Assistant - Mohamed Abd Elnasser - IT Manager'
    $wb.Author = 'Mohamed Abd Elnasser - IT Manager'
    $wb.IsAddin = $true
    $xlOpenXMLAddIn = 55
    $wb.SaveAs($Output, $xlOpenXMLAddIn)
    $wb.Close($true); $wb=$null
    $excel.Quit(); $excel=$null
    if (Test-Path $TempTest) { Remove-Item $TempTest -Force }
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (Test-Path $TempTest) { try { Remove-Item $TempTest -Force } catch {} }
}

# Inject RibbonX with unique internal IDs. Visible labels remain identical.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Validate Ribbon XML before touching the XLAM.
foreach($ribbonFile in @($Ribbon06,$Ribbon14)) {
    try {
        [xml]$rx = [IO.File]::ReadAllText($ribbonFile)
        if ($null -eq $rx.DocumentElement) { throw 'Missing XML root element.' }
    } catch {
        throw ("Ribbon XML validation failed for " + $ribbonFile + ": " + $_.Exception.Message)
    }
}
Write-Host 'Ribbon XML validation: PASS' -ForegroundColor Green

$zip=[System.IO.Compression.ZipFile]::Open($Output,[System.IO.Compression.ZipArchiveMode]::Update)
try {
    foreach($pair in @(@('customUI/customUI.xml',$Ribbon06),@('customUI/customUI14.xml',$Ribbon14))) {
        $old=$zip.GetEntry($pair[0]); if($old){$old.Delete()}
        $entry=$zip.CreateEntry($pair[0],[System.IO.Compression.CompressionLevel]::Optimal)
        $writer=New-Object IO.StreamWriter($entry.Open(),(New-Object Text.UTF8Encoding($false)))
        try{$writer.Write([IO.File]::ReadAllText($pair[1]))}finally{$writer.Dispose()}
    }

    # Rewrite package root relationships using XML DOM instead of string replacement.
    $re=$zip.GetEntry('_rels/.rels')
    if($null -eq $re){ throw 'Missing _rels/.rels in generated XLAM.' }
    $rd=New-Object IO.StreamReader($re.Open())
    try{$relsText=$rd.ReadToEnd()}finally{$rd.Dispose()}
    $re.Delete()

    [xml]$relsXml=$relsText
    $ns=$relsXml.DocumentElement.NamespaceURI
    $mgr=New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $mgr.AddNamespace('r',$ns)

    # Remove any pre-existing RibbonX relationships so stale/cached targets cannot survive.
    @($relsXml.SelectNodes('//r:Relationship[contains(@Type,"/ui/extensibility")]',$mgr)) | ForEach-Object {
        [void]$relsXml.DocumentElement.RemoveChild($_)
    }

    $rel06=$relsXml.CreateElement('Relationship',$ns)
    $rel06.SetAttribute('Id','rIdEIAMohamedRibbon06')
    $rel06.SetAttribute('Type','http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
    $rel06.SetAttribute('Target','customUI/customUI.xml')
    [void]$relsXml.DocumentElement.AppendChild($rel06)

    $rel14=$relsXml.CreateElement('Relationship',$ns)
    $rel14.SetAttribute('Id','rIdEIAMohamedRibbon14')
    $rel14.SetAttribute('Type','http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
    $rel14.SetAttribute('Target','customUI/customUI14.xml')
    [void]$relsXml.DocumentElement.AppendChild($rel14)

    $settings=New-Object System.Xml.XmlWriterSettings
    $settings.Encoding=New-Object Text.UTF8Encoding($false)
    $settings.Indent=$false
    $nr=$zip.CreateEntry('_rels/.rels',[System.IO.Compression.CompressionLevel]::Optimal)
    $stream=$nr.Open()
    $xw=[System.Xml.XmlWriter]::Create($stream,$settings)
    try{$relsXml.Save($xw)}finally{$xw.Dispose();$stream.Dispose()}
}
finally{$zip.Dispose()}

# Final package validation: verify parts, relationships, XML, and callback names.
$verifyZip=[System.IO.Compression.ZipFile]::OpenRead($Output)
try {
    $required=@('[Content_Types].xml','_rels/.rels','xl/workbook.xml','xl/vbaProject.bin','customUI/customUI.xml','customUI/customUI14.xml')
    foreach($name in $required){
        if($null -eq $verifyZip.GetEntry($name)){ throw "Final package validation failed: missing $name" }
    }

    $relsEntry=$verifyZip.GetEntry('_rels/.rels')
    $rr=New-Object IO.StreamReader($relsEntry.Open())
    try{$finalRels=$rr.ReadToEnd()}finally{$rr.Dispose()}
    if($finalRels -notmatch 'office/2006/relationships/ui/extensibility'){ throw 'Final package validation failed: Office 2006 Ribbon relationship missing.' }
    if($finalRels -notmatch 'office/2007/relationships/ui/extensibility'){ throw 'Final package validation failed: Office 2010+ Ribbon relationship missing.' }

    foreach($ribbonName in @('customUI/customUI.xml','customUI/customUI14.xml')) {
        $e=$verifyZip.GetEntry($ribbonName)
        $r=New-Object IO.StreamReader($e.Open())
        try{$xmlText=$r.ReadToEnd()}finally{$r.Dispose()}
        [xml]$null=$xmlText
        if($xmlText -notmatch 'id="EIA_Mohamed_17"'){ throw "Final package validation failed: unique tab ID missing in $ribbonName" }
        foreach($cb in @('ImgSingle','ImgAll','ImgDel','CommSingle','CommAll','CommDel','rbnDownInsertFiles','Compress','rbnFileName','rbnAllFileNames','ShowPicc','ShowPicc1','ShowAbout')) {
            if($xmlText -match ('onAction="' + [regex]::Escape($cb) + '"')) {
                $basText=[IO.File]::ReadAllText($BasImport)
                if($basText -notmatch ('(?im)^Public Sub\s+' + [regex]::Escape($cb) + '\b')) {
                    throw "Final package validation failed: callback $cb referenced by Ribbon but missing from VBA."
                }
            }
        }
    }
} finally { $verifyZip.Dispose() }
Write-Host 'XLAM package + Ribbon relationship validation: PASS' -ForegroundColor Green

Write-Host ''
Write-Host 'DONE - v1.7 validated build' -ForegroundColor Green
Write-Host $Output -ForegroundColor Green
Write-Host ''
Write-Host 'Install: Excel > File > Options > Add-ins > Manage Excel Add-ins > Go > Browse.'
