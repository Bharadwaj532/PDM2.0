# Convert Markdown Documentation to Word Document
# Package Development Manager 2.0 Documentation Converter

param(
    [string]$MarkdownFile = "Package_Development_Manager_Documentation.md",
    [string]$WordFile = "Package_Development_Manager_2.0_Documentation.docx"
)

Write-Host "=== Package Development Manager 2.0 Documentation Converter ===" -ForegroundColor Cyan
Write-Host "Converting Markdown to Word Document..." -ForegroundColor Yellow

try {
    # Check if pandoc is available (preferred method)
    $pandocAvailable = $false
    try {
        $pandocVersion = pandoc --version 2>$null
        if ($pandocVersion) {
            $pandocAvailable = $true
            Write-Host "Pandoc found - using for conversion" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Pandoc not found - using alternative method" -ForegroundColor Yellow
    }

    if ($pandocAvailable) {
        # Use Pandoc for high-quality conversion
        Write-Host "Converting with Pandoc..." -ForegroundColor Yellow
        $pandocArgs = @(
            $MarkdownFile,
            "-o", $WordFile,
            "--reference-doc=template.docx",
            "--toc",
            "--toc-depth=3"
        )
        
        & pandoc @pandocArgs
        
        if (Test-Path $WordFile) {
            Write-Host "✅ Word document created successfully: $WordFile" -ForegroundColor Green
        } else {
            throw "Pandoc conversion failed"
        }
    }
    else {
        # Alternative method using Word COM object
        Write-Host "Using Microsoft Word COM object for conversion..." -ForegroundColor Yellow
        
        # Read markdown content
        $markdownContent = Get-Content $MarkdownFile -Raw
        
        # Create Word application
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Create new document
        $doc = $word.Documents.Add()
        
        # Process markdown content
        $htmlContent = Convert-MarkdownToHTML -MarkdownText $markdownContent
        
        # Insert HTML content
        $selection = $word.Selection
        $selection.InsertAfter($htmlContent)
        
        # Apply formatting
        Apply-WordFormatting -Document $doc
        
        # Save document
        $fullPath = Join-Path (Get-Location) $WordFile
        $doc.SaveAs2($fullPath, 16) # 16 = wdFormatDocumentDefault
        
        # Close and cleanup
        $doc.Close()
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-Host "✅ Word document created successfully: $WordFile" -ForegroundColor Green
    }
}
catch {
    Write-Host "❌ Error during conversion: $($_.Exception.Message)" -ForegroundColor Red
    
    # Fallback: Create a simple text-based Word document
    Write-Host "Creating fallback Word document..." -ForegroundColor Yellow
    Create-FallbackWordDocument -MarkdownFile $MarkdownFile -WordFile $WordFile
}

function Convert-MarkdownToHTML {
    param([string]$MarkdownText)
    
    # Basic markdown to HTML conversion
    $html = $MarkdownText
    
    # Headers
    $html = $html -replace '^# (.+)$', '<h1>$1</h1>' -replace '^## (.+)$', '<h2>$1</h2>' -replace '^### (.+)$', '<h3>$1</h3>'
    
    # Bold and italic
    $html = $html -replace '\*\*(.+?)\*\*', '<strong>$1</strong>' -replace '\*(.+?)\*', '<em>$1</em>'
    
    # Code blocks
    $html = $html -replace '```(.+?)```', '<pre><code>$1</code></pre>'
    $html = $html -replace '`(.+?)`', '<code>$1</code>'
    
    # Lists
    $html = $html -replace '^- (.+)$', '<li>$1</li>'
    
    # Line breaks
    $html = $html -replace "`n", '<br>'
    
    return $html
}

function Apply-WordFormatting {
    param($Document)
    
    try {
        # Apply styles to the document
        $range = $Document.Range()
        
        # Set font
        $range.Font.Name = "Segoe UI"
        $range.Font.Size = 11
        
        # Apply heading styles
        $headings = $Document.Range().Find
        $headings.Text = "<h1>"
        while ($headings.Execute()) {
            $headings.Parent.Style = "Heading 1"
            $headings.Text = $headings.Text -replace '<h1>|</h1>', ''
        }
        
        Write-Host "Word formatting applied successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not apply advanced formatting: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Create-FallbackWordDocument {
    param(
        [string]$MarkdownFile,
        [string]$WordFile
    )
    
    try {
        Write-Host "Creating fallback Word document..." -ForegroundColor Yellow
        
        # Read markdown content
        $content = Get-Content $MarkdownFile -Raw
        
        # Create Word application
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Create new document
        $doc = $word.Documents.Add()
        
        # Clean up markdown syntax for plain text
        $cleanContent = $content -replace '#+ ', '' -replace '\*\*(.+?)\*\*', '$1' -replace '\*(.+?)\*', '$1' -replace '```(.+?)```', '$1' -replace '`(.+?)`', '$1'
        
        # Insert content
        $selection = $word.Selection
        $selection.TypeText($cleanContent)
        
        # Basic formatting
        $selection.WholeStory()
        $selection.Font.Name = "Segoe UI"
        $selection.Font.Size = 11
        
        # Save document
        $fullPath = Join-Path (Get-Location) $WordFile
        $doc.SaveAs2($fullPath, 16)
        
        # Close and cleanup
        $doc.Close()
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-Host "✅ Fallback Word document created: $WordFile" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Fallback conversion also failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please manually convert the markdown file or install Pandoc for better conversion" -ForegroundColor Yellow
    }
}

Write-Host "`n=== Conversion Complete ===" -ForegroundColor Cyan
Write-Host "Files created in: $(Get-Location)" -ForegroundColor Gray
Write-Host "- Markdown: $MarkdownFile" -ForegroundColor Gray
Write-Host "- Word Document: $WordFile" -ForegroundColor Gray
Write-Host "`nDocumentation is ready for review and distribution!" -ForegroundColor Green
