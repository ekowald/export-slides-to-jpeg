# Converts pptx to jpeg. Requires slides.pptx of each slides.md file.
# Run script using argument that points to module folder
# .\export-slides-to-jpeg.ps1 <path\to\module>

if ( $args.Count -eq 0 ) 
{
    Write-Host "Must supply Module folder as argument.`n .\export-slides-to-jpeg.ps1 <path\to\module>" -ForegroundColor Red
    return
}

$msppt = New-Object -ComObject powerpoint.application

$moduleFolder = $args[0]

$lessons = get-childitem "$moduleFolder\lessons"

$moduleName = $moduleFolder | Split-Path -Leaf
Write-Host "Converting '$moduleName' slides`n" -ForegroundColor DarkBlue

foreach ( $lesson in $lessons ) 
{   

    if ( ( Test-Path "$lesson\slides.pptx" ) -and ( ( ( Get-Item "$lesson\slides.pptx" ).length ) -gt ( 1KB ) ) )
    {   
        $lessonName = $lesson.Name     
        Write-Host "Creating /slides/ directory in $lessonName`n" -ForegroundColor Green
        $null = New-Item -Path $lesson.fullname -Name slides -ItemType Directory -Force
        
        $pres = $msppt.presentations.open( "$lesson\slides.pptx", 2, $True, $False )
        Write-Host "Converting $lessonName\slides.pptx to jpeg`n" -ForegroundColor Yellow

        foreach ( $slide in $pres.slides )
        {
            $slide.Export( "$lesson\slides\" + $slide.Name + ".jpeg", "JPG" )
        }
    }

    # In progress attempt at re-naming jpeg files
    get-childitem -Path $lesson\slides\* -Include *.jpeg -recurse | foreach-object
    {
    $tmp = $_.name
    $tmp = $tmp -split ' '
    $noNum = $true
    $tmp = foreach ($s in $tmp) 
    {
        if($noNum)
        {
            if($s -match "\b\d\b"){"0$s"; $noNum = $false}
            elseif($s -match "\b\d\d\b"){"$s"; $noNum = $false}
            else{$s}
        }
        else {$s}
    }
    $tmp = $tmp -join ' '
    rename-item $_ $tmp
    }

    Get-ChildItem *.jpeg -recurse | Rename-Item -NewName {$_.Name -replace "Slide ", "" } # End in progress attempt
}

$lessonNames = $lessons.Name | Out-String

Write-Host "Finished converting pptx to jpeg for:"
Write-Host $lessonNames -ForegroundColor Blue

$msppt.Quit()