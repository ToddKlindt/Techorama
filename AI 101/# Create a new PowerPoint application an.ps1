# Create a new PowerPoint application and presentation
$PowerPoint = New-Object -ComObject PowerPoint.Application
$PowerPoint.Visible = $true
$Presentation = $PowerPoint.Presentations.Add()

# Function to add a slide with title, content, speaker notes, and image
function Add-Slide {
    param (
        [int]$SlideIndex,
        [string]$Title,
        [string]$Content,
        [string]$Notes,
        [string]$ImagePath
    )
    $slide = $Presentation.Slides.Add($SlideIndex, 2)
    $slide.Shapes.Title.TextFrame.TextRange.Text = $Title
    $slide.Shapes[2].TextFrame.TextRange.Text = $Content
    $slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = $Notes

    # Add an image to the slide
    if (Test-Path $ImagePath) {
        $image = $slide.Shapes.AddPicture($ImagePath, "MsoTriState::msoFalse", "MsoTriState::msoCTrue", 480, 50, 300, 200)
    } else {
        Write-Host "Image path not found: $ImagePath"
    }
}

# Add each slide with an image path parameter
# Example usage (You need to replace 'Path\To\Image.jpg' with actual image paths)
Add-Slide -SlideIndex 1 -Title "AI 101: Unveiling the Mechanics of Artificial Intelligence" -Content "Presenter's Name`nDate: [Date]`nEvent: Techorama" -Notes "Introduce yourself and the topic. Mention your excitement about discussing AI at Techorama." -ImagePath "C:\Path\To\TitleImage.jpg"

# Repeat for other slides, ensuring each slide has a relevant image path
# Make sure to add correct paths for each image

# Save and close operations as before
$Presentation.SaveAs(".\Code\AI_101_Presentation.pptx")
break
$PowerPoint.Quit()

# Cleanup COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
