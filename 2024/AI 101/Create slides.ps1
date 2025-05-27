# Create a new PowerPoint application and presentation
$PowerPoint = New-Object -ComObject PowerPoint.Application
$PowerPoint.Visible = $true
$Presentation = $PowerPoint.Presentations.Add()

# Function to add a slide with title and content
function Add-Slide {
    param (
        [int]$SlideIndex,
        [string]$Title,
        [string]$Content
    )
    $slide = $Presentation.Slides.Add($SlideIndex, 1)
    $slide.Shapes.Title.TextFrame.TextRange.Text = $Title
    $slide.Shapes[2].TextFrame.TextRange.Text = $Content
}

# Slide 1: Title Slide
Add-Slide -SlideIndex 1 -Title "AI 101: Unveiling the Mechanics of Artificial Intelligence" -Content "Presenter's Name`nDate: [Date]`nEvent: Techorama"

# Slide 2: Introduction
Add-Slide -SlideIndex 2 -Title "Introduction" -Content "Session Overview`nObjectives`nWhy AI Matters"

# Slide 3: A Brief History of AI
Add-Slide -SlideIndex 3 -Title "A Brief History of AI" -Content "Early Concepts and Theories`nKey Milestones in AI Development`nFrom Theory to Practice"

# Slides 4-6: Understanding Word Vectors
Add-Slide -SlideIndex 4 -Title "What are Word Vectors?" -Content "Definition and Basic Concept"
Add-Slide -SlideIndex 5 -Title "Creating Word Vectors" -Content "Process and Techniques (e.g., Word2Vec)"
Add-Slide -SlideIndex 6 -Title "Word Vectors in Use" -Content "Examples in Natural Language Processing"

# Slides 7-9: Dive into Transformers
Add-Slide -SlideIndex 7 -Title "The Architecture of Transformers" -Content "Basic Structure and Components"
Add-Slide -SlideIndex 8 -Title "Understanding Self-Attention Mechanisms" -Content "How Self-Attention Works"
Add-Slide -SlideIndex 9 -Title "Applications of Transformers" -Content "Examples in Real-World Applications"

# Slide 10: Current Trends and Future Directions
Add-Slide -SlideIndex 10 -Title "Current Trends and Future Directions" -Content "Latest Trends in AI`nSpeculative Future Developments"

# Slide 11: Q&A Session
Add-Slide -SlideIndex 11 -Title "Questions & Answers" -Content "Any questions? Let's discuss!"

# Slide 12: Conclusion
Add-Slide -SlideIndex 12 -Title "Conclusion" -Content "Summary of Key Points`nThank You and Contact Information"


# Save the presentation
$Presentation.SaveAs(".\AI_101_Presentation.pptx")
break
$PowerPoint.Quit()

# Cleanup COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
