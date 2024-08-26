# Define paths
$orgFilePath = "C:\Users\sannomiya\Documents\Learning\ropiron.github.io\memo.org"
$pptxFilePath = "C:\Users\sannomiya\Documents\Learning\ropiron.github.io\output.pptx"

# Read the Org-mode file content with UTF-8 encoding
$orgContent = Get-Content $orgFilePath -Encoding UTF8

# Create a new PowerPoint application and presentation
$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $pptApp.Presentations.Add()

# Initialize variables
$slide = $null
$title = $null
$isFirstSlide = $true

# Parse Org-mode content
foreach ($line in $orgContent) {
    if ($line.StartsWith("*")) {
        $title = $line.TrimStart("* ").Trim()
        
        if ($isFirstSlide) {
            # Create the first slide with title only
            $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 1)
            $slide.Shapes.Title.TextFrame.TextRange.Text = $title
            $isFirstSlide = $false
        } else {
            # Create subsequent slides with title and content layout
            $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 2)
            $slide.Shapes.Title.TextFrame.TextRange.Text = $title
            # Clear any existing text in the content placeholder
            $slide.Shapes[2].TextFrame.TextRange.Text = ""
        }
    } elseif ($line.TrimStart().StartsWith("-")) {
        # Determine the indent level based on the number of leading spaces
        $indentLevel = ($line.Length - $line.TrimStart().Length) / 2
        
        # Add bullet points to the current slide
        $bulletPoint = $line.TrimStart("- ").Trim()
        $textBox = $slide.Shapes[2].TextFrame.TextRange
        # Initialize the content only if it's the first bullet point
        if ($textBox.Text -eq "") {
            $textBox.Text = $bulletPoint
        } else {
            $textBox.Text = $textBox.Text + "`n" + $bulletPoint
        }
        
        # Set the indent level for the current bullet point
        $textBox.ParagraphFormat.IndentLevel = $indentLevel
    }
}

# Save the presentation
$presentation.SaveAs($pptxFilePath)
$presentation.Close()
$pptApp.Quit()

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
Remove-Variable pptApp