Set-Location $PSScriptRoot
cls
<#
  First things first: Get iTextSharp e.g. from Nuget:
  https://www.nuget.org/packages/iTextSharp/
  (Rename file from .nupkg to .zip and unpack the .dll)
#>

# We'll first run this example and watch the resulting PDF
# Then we'll have a look under the hood:

function Create-FirstPDF
{
  Begin
  {
    Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
    $file = $(Join-Path $pwd "MyNewPDF.pdf")  
    $doc  = New-Object iTextSharp.text.Document
    $stream = [IO.File]::OpenWrite($file)
    $writer = [itextsharp.text.pdf.PdfWriter]::GetInstance($doc, $stream)
    
  }
  Process
  {
    # Metadata for the PDF
    [void]$doc.SetPageSize([iTextSharp.text.PageSize]::A4)
    [void]$doc.SetMargins(20,20,20,20)
    [void]$doc.AddAuthor($env:USERNAME)
    [void]$doc.AddSubject("My first PDF document with PowerShell and iTextSharp")

    # Open the document object
    $doc.Open()
    
    # Set fonts
    $redTimeFont28 = [iTextSharp.text.FontFactory]::GetFont("TIMES", 28, [iTextSharp.text.BaseColor]::RED)
    $blueBoldHelveticaFont28 = [iTextSharp.text.FontFactory]::GetFont("HELVETICA_BOLD", 28, [iTextSharp.text.BaseColor]::BLUE)
    $BoldFont = [iTextSharp.text.FontFactory]::GetFont("HELVETICA_BOLD", 18, [iTextSharp.text.BaseColor]::DARK_GRAY)
    
    # Write the first paragraph
    $p = New-Object iTextSharp.text.Paragraph
    [void]$p.Add([iTextSharp.text.Paragraph]::new("Hello PSConfEU 2019 !!!", $blueBoldHelveticaFont28))   
    [void]$doc.Add($p)
    
    # Chunk is the smallest part of text that can be added to a document  
    $doc.Add([iTextSharp.text.Chunk]("")) 
        
    # Example table with XML
    [xml]$x = Get-Content -Path $(Join-Path $pwd "recipe.xml")
    
    $table = New-Object iTextSharp.text.pdf.PDFPTable(2)
    $table.TotalWidth = 400
    $table.DefaultCell.Padding = 2
          
    $header = New-Object iTextSharp.text.pdf.PdfPCell([iTextSharp.text.Paragraph]::new($x.recipe.title, $redTimeFont28))   
    $header.Colspan = 2
    $header.Padding = 5
    $header.HorizontalAlignment = 1 # 0=Left, 1=Centre, 2=Right
    $table.AddCell($header)
    $table.AddCell([iTextSharp.text.Paragraph]::new("Preparation Time:", $BoldFont))
    $table.AddCell($x.recipe.prep_time)
    $table.AddCell([iTextSharp.text.Paragraph]::new("Cook Time:", $BoldFont))
    $table.AddCell($x.recipe.cook_time)

    $ingredients = New-Object iTextSharp.text.pdf.PdfPCell([iTextSharp.text.Paragraph]::new("Ingredients:", $BoldFont))     
    $ingredients.Padding = 5
    $ingredients.HorizontalAlignment = 0
    $table.AddCell($ingredients)
    
    $nested = New-Object iTextSharp.text.pdf.PDFPTable(3) 
    $ne
    
    for ($i=0; $i -le ($x.recipe.ingredient.length -1); $i++){
    
      $nested.AddCell($x.recipe.ingredient[$i].amount)
      $nested.AddCell($x.recipe.ingredient[$i].unit)
      $nested.AddCell($x.recipe.ingredient[$i]."#text")
    }
    
    $nesthousing = New-Object iTextSharp.text.pdf.PdfPCell($nested)
    
    $table.AddCell($nesthousing)
    
    $instruction = New-Object iTextSharp.text.pdf.PdfPCell([iTextSharp.text.Paragraph]::new("Instructions:", $BoldFont))   
    $instruction.Colspan = 2
    $instruction.Padding = 5
    $instruction.HorizontalAlignment = 0 
    $table.AddCell($instruction)
    
    $theInstructions = $x.recipe.instructions.step | out-string
    $myInstruction = New-Object iTextSharp.text.pdf.PdfPCell($theInstructions)   
    $myInstruction.Colspan = 2
    $myInstruction.Padding = 5
    $myInstruction.HorizontalAlignment = 0 
    $table.AddCell($myInstruction)
    
    # Insert an image
    $img = [iTextSharp.text.Image]::GetInstance($(Join-Path "$($pwd)" "PowerShell.png"))
    $imgCell = New-Object iTextSharp.text.pdf.PdfPCell($img)
    $imgCell.Colspan = 2
    $imgCell.Border = 0
    $imgCell.Padding = 10
    $table.AddCell($imgCell)
    
    
    # Write the table to the PDF document
    $doc.Add($table)
    
    # Write the graph:
    
    $cb = $writer.DirectContent
      
    # Start Point
    $cb.MoveTo(200, 10)

    # Control Point 1, Control Point 2, End Point
    $cb.CurveTo(150, 30, 450, 70, 350, 150)
    $cb.Stroke()

    $cb.MoveTo(200, 10)
    $cb.LineTo(150, 30)
    $cb.Stroke()
    $cb.MoveTo(450, 70)
    $cb.LineTo(350, 150)
    $cb.Stroke()
    $cb.Circle(450, 70, 1)
    $cb.Stroke()
    $cb.Circle(150, 30, 1)
    $cb.Stroke()

    $cb.SetColorStroke([iTextSharp.text.BaseColor]::GREEN)
    
    # start pint (x,y)
    $cb.MoveTo(200, 10)
    $cb.CurveTo(150, 30, 550, 100, 350, 150)
    $cb.Stroke()
    
    $cb.MoveTo(550, 100)
    $cb.LineTo(350, 150)
    $cb.Stroke()
    
    $cb.Circle(550, 100, 1)
    $cb.Stroke()
    
    $cb.SetColorStroke([iTextSharp.text.BaseColor]::LIGHT_GRAY)
    
    $cb.Arc(350, 70, 550, 130, 270, 90)
    $cb.SetLineDash(3, 3)
    $cb.Stroke()
    
    $cb.SetLineDash(0)
    $cb.MoveTo(550, 100)
    $cb.LineTo(535, 95)
    $cb.Stroke()
    $cb.MoveTo(550, 100)
    $cb.LineTo(552, 88)
    $cb.Stroke()

  }
  End
  {
    $doc.Dispose()
    $stream.Close()
  }
}

Create-FirstPDF 
Invoke-Expression $(Join-Path $pwd "MyNewPDF.pdf") 
