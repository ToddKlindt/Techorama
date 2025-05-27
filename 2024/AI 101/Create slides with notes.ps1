# Create a new PowerPoint application and presentation
$PowerPoint = New-Object -ComObject PowerPoint.Application
$PowerPoint.Visible = $true
$Presentation = $PowerPoint.Presentations.Add()

# Function to add a slide with title, content, and speaker notes
function Add-Slide {
    param (
        [int]$SlideIndex,
        [string]$Title,
        [string]$Content,
        [string]$Notes
    )
    $slide = $Presentation.Slides.Add($SlideIndex, 1)
    $slide.Shapes.Title.TextFrame.TextRange.Text = $Title
    $slide.Shapes[2].TextFrame.TextRange.Text = $Content
    $slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = $Notes
}

# Slide 1: Title Slide
Add-Slide -SlideIndex 1 -Title "AI 101: Unveiling the Mechanics of Artificial Intelligence" -Content "Presenter's Name`nDate: [Date]`nEvent: Techorama" -Notes "Introduce yourself and the topic. Mention your excitement about discussing AI at Techorama."

# Slide 2: Introduction
Add-Slide -SlideIndex 2 -Title "Introduction" -Content "Session Overview`nObjectives`nWhy AI Matters" -Notes "Briefly explain what will be covered and why understanding AI is crucial today."

# Slide 3: A Brief History of AI
Add-Slide -SlideIndex 3 -Title "A Brief History of AI" -Content "Early Concepts and Theories`nKey Milestones in AI Development`nFrom Theory to Practice" -Notes "Highlight significant milestones in AI, showing its evolution and increasing impact on technology."

# Slides 4-6: Understanding Word Vectors
Add-Slide -SlideIndex 4 -Title "What are Word Vectors?" -Content "Definition and Basic Concept" -Notes "Explain the concept of word vectors with a simple example or analogy."
Add-Slide -SlideIndex 5 -Title "Creating Word Vectors" -Content "Process and Techniques (e.g., Word2Vec)" -Notes "Discuss the technical process of creating word vectors, possibly mentioning different models like Word2Vec."
Add-Slide -SlideIndex 6 -Title "Word Vectors in Use" -Content "Examples in Natural Language Processing" -Notes "Provide real-world applications to demonstrate how word vectors are used in AI solutions."

# Slides 7-9: Dive into Transformers
Add-Slide -SlideIndex 7 -Title "The Architecture of Transformers" -Content "Basic Structure and Components" -Notes "Describe the architecture, focusing on why transformers are effective for processing language."
Add-Slide -SlideIndex 8 -Title "Understanding Self-Attention Mechanisms" -Content "How Self-Attention Works" -Notes "Go into detail about the self-attention mechanism, explaining its significance."
Add-Slide -SlideIndex 9 -Title "Applications of Transformers" -Content "Examples in Real-World Applications" -Notes "List and explain some key applications of transformers, such as in natural language understanding."

# Slide 10: Current Trends and Future Directions
Add-Slide -SlideIndex 10 -Title "Current Trends and Future Directions" -Content "Latest Trends in AI`nSpeculative Future Developments" -Notes "Discuss emerging trends and potential future developments in AI, encouraging excitement and curiosity."

# Slide 11: Q&A Session
Add-Slide -SlideIndex 11 -Title "Questions & Answers" -Content "Any questions? Let's discuss!" -Notes "Encourage audience interaction, prepare to engage with their queries."

# Slide 12: Conclusion
Add-Slide -SlideIndex 12 -Title "Conclusion" -Content "Summary of Key Points`nThank You and Contact Information" -Notes "Summarize the main points discussed, thank the audience for their participation, and provide contact details for further interaction."

# Save the presentation
$Presentation.SaveAs(".\AI_101_Presentation.pptx")
break
$PowerPoint.Quit()

# Cleanup COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
