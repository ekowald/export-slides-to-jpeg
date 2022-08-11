# Check that argument was passed and that it is a directory
# Exit if not
function checkArgs ( $module )
{
    if ( !( Test-Path -Path $module ) )
    {
        Write-Host "Module Folder not found." -ForegroundColor Red
        exit
    }
    
    else
    {
        $moduleName = $module | Split-Path -Leaf
        Write-Host "Converting '$moduleName' slides to jpeg images.`n" -ForegroundColor Yellow
        createSlidesDirectory ( $module )
    }
}

# Create a new directory for the jpeg images
function createSlidesDirectory ( $module )
{
    $lessons = Get-ChildItem -Path "$module\lessons"

    foreach ( $lesson in $lessons )
    {
        $lessonName = $lesson.Name

        if ( ( Test-Path "$lesson\slides.pptx" ) -and ( ( ( Get-Item "$lesson\slides.pptx" ).length ) -gt ( 1KB ) ) )
        {

            if ( !( Test-Path "$lesson\slides" ) )
            {
                Write-Host "Creating directory for $lessonName slides." -ForegroundColor Yellow
                $null = New-Item -Path $lesson.fullname -Name slides -ItemType Directory -Force
                Write-Host "Directory creation successful.`n" -ForegroundColor Green
                convertSlides ( $lesson )
            }

            else
            {
                $dir = "$lesson\slides"
                Remove-Item $dir -Force -Recurse
                
                Write-Host "Creating directory for $lessonName slides." -ForegroundColor Yellow
                $null = New-Item -Path $lesson.fullname -Name slides -ItemType Directory -Force
                Write-Host "Directory creation successful.`n" -ForegroundColor Green
                convertSlides ( $lesson )
            }
        }
    }
}

# Convert the slides to jpeg images
function convertSlides ( $lesson )
{
    $lessonName = $lesson.Name
    $msppt = New-Object -ComObject powerpoint.application
    $pres = $msppt.presentations.open( "$lesson\slides.pptx", 2, $True, $False )

    Write-Host "Converting $lessonName\slides.pptx to jpeg" -ForegroundColor Yellow

    foreach ( $slide in $pres.slides )
    {
        $slide.Export( "$lesson\slides\" + $slide.Name + ".jpeg", "JPG" )
    }

    Write-Host "Finished Converting $lessonName\slides.pptx to jpeg`n" -ForegroundColor Green
    $msppt.Quit()


    $slidesDir = "$lesson\slides"
    changeNames ( $slidesDir )

}

# Change the names of the jpeg images to match the slide number
function changeNames ( $slidesDir )
{
    $slides = Get-ChildItem -Path $slidesDir
    Write-Host "Converting JPEG filenames" -ForegroundColor Yellow
    
    foreach ( $slide in $slides )
    {                
        $tmp = $slide.Name
        $tmp = $tmp -split ' '
        $noNum = $true
        $tmp = foreach ( $s in $tmp )
        {
            if ( $noNum )
            {
                if ( $s -match "\b\d\b" )
                {
                    "0$s"
                    $noNum = $false
                }

                elseif ( $s -match "\b\d\d\b" )
                {
                    "$s"
                    $noNum = $false
                }

                else 
                {
                    $s
                }
            }
        }
        $tmp = $tmp -join ' '
        $tmp = $tmp -replace 'Slide ', ''
        Rename-Item -Path $slide -NewName $tmp      
    }

    Write-Host "Finished Converting JPEG filenames`n" -ForegroundColor Green
}

# Begin Script
# If no args supplied, exit.
if ( $args.Count -lt 1 )
{
    Write-Host "Must supply Module Folder path. .\export-slides-to-jpeg.ps1 <Module\Folder\Path>" -ForegroundColor Red
    exit
}

$moduleFolderPath = $args[0]
checkArgs( $moduleFolderPath )