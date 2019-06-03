Set-Location $PSScriptRoot

function Get-PDFInfo
{
  <#
      .Synopsis
      Get Infos from a PDF file
      .DESCRIPTION
      Use iTextSharp PDFReader class to get infos from a given PDF
      .EXAMPLE
      Get-PDFInfo -File YourFile.pdf
  #>
  [CmdletBinding()]
  Param
  (
    # Insert filename
    [String]
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
      Position=0)
    ]
    $Name,
         
    [switch]$Javascript
  )
 
  Begin
  {
    Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
    $reader  = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $(Join-Path $pwd $Name)
  }
  Process
  {
    # Output of PDFReader Info is a hashtable:
    $info = $reader.Info
     
    # Change DateTime to human-readable format and add both to the hashtable again
    $ModDate = (Select-String -InputObject $info.ModDate -Pattern "\d{14}").Matches.Value
    $CreationDate = (Select-String -InputObject $info.CreationDate -Pattern "\d{14}").Matches.Value
    $info.ModDate = [datetime]::ParseExact($ModDate,"yyyyMMddHHmmss", $null)
    $info.CreationDate = [datetime]::ParseExact($CreationDate,"yyyyMMddHHmmss",$null)
     
    # Add more keys to the hashtable:
    $info += @{NumberOfPages = $reader.NumberOfPages}
    $info += @{FileLength = $reader.FileLength}
    $info += @{PdfVersion = $reader.PdfVersion}
    $info += @{IsEncrypted = $reader.IsEncrypted()}
     
    $info += @{PageSize = $reader.GetPageSize(1)}
    
    # Maybe the PDF contains JavaScript?
    if ($Javascript)
    {
      $info += @{JavaScript = $reader.JavaScript}
    }    
  }
  End
  {
    $info.GetEnumerator() | Sort-Object -Property name
  }
}
 
 # Get-PDFInfo demo:
<#

    Get-PDFInfo MyNewPDF.pdf
    Get-PDFInfo Christian_Imhorst.pdf

#>

function Set-PDFMetadata
{
  <#
      .Synopsis
      Set or change metadata in your PDF file
      .DESCRIPTION
      Set or change metadata in your PDF file with iTextSharp
      .EXAMPLE
      Set-PDFMetadata -File Input.pdf -Output Output_neu.pdf -Metadata @{"Author" = "Christian Imhorst"}
      .EXAMPLE
      Set-PDFMetadata -File Input.pdf -Output Output_neu.pdf -Metadata @{"Author" = "Christian Imhorst"; "Creator" = "PowerShell"; "Conference" = "PSConfEU2019"; "Hashtag" = "#PSConfEU2019"}
  #>
  Param
  (
    # Insert filename
    [String]
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
        Position=0)
    ]
    $File,
    # Name of the output file    
    [String]
    [Parameter(
        Mandatory,
        Position=1)
    ]
    $Output,
    
    # Hashtable with your metadata: @{"Author" = "Christian Imhorst"}
    [hashtable]
    [Parameter(
        Mandatory,
        Position=2)
    ]
    $Metadata
  )

  Begin
  {
    Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
    $reader  = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $(Join-Path $pwd $File)
    $fs = [System.IO.FileStream]::new($(Join-Path $pwd $Output), [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, $fs)  
  }
  Process
  {
    $info = $reader.Info
    foreach ($key in $Metadata.Keys)
    {
      if ($info.ContainsKey($key))
      {
        $info.Remove($key)
      }
      $info.Add($key, $Metadata[$key])
    }
    
    $stamper.MoreInfo = $info
    $stamper.Dispose()
  }
  End
  {
    $fs.Dispose()
    $reader.Dispose()
  }
}

# Set-PDFMetadata demo:
<#
    Set-PDFMetadata -File Christian_Imhorst.pdf -Output Christian_Imhorst_new.pdf -Metadata @{"Author" = "Christian Imhorst"}
    Get-PDFInfo Christian_Imhorst_new.pdf
    Invoke-Expression $(Join-Path $pwd "Christian_Imhorst_new.pdf")
  
    Set-PDFMetadata -File Christian_Imhorst.pdf -Output Christian_Imhorst_new.pdf -Metadata @{"Author" = "Christian Imhorst"; "Creator" = "PowerShell"; "Conference" = "PSConfEU2019"; "Hashtag" = "#PSConfEU2019"}
    Get-PDFInfo Christian_Imhorst_new.pdf
    Invoke-Expression $(Join-Path $pwd "Christian_Imhorst_new.pdf")
#>
 
function Export-TextFromPDF
{
  <#
      .Synopsis
      Export Text from a PDF file
      .DESCRIPTION
      Use the iTextSharp parser PdfTextExtractor to get only the text from a given PDF
      .EXAMPLE
      Export-TextFromPDF -File YourFile.pdf
  #>
  [CmdletBinding()]
  Param
  (
    # Insert filename
    [String]
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
      Position=0)
    ]
    $Name
  )
   
  Begin
  {
    Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
    $reader  = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $(Join-Path $pwd $Name)

  }
  Process
  {
    for ($i = 1; $i -lt $reader.NumberOfPages; $i++) 
    {
      $text = $text + [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i)
    }
  }
  End
  {
    $text
  }
   
}

# Export-TextFromPDF demo:
<#
    Export-TextFromPDF Christian_Imhorst.pdf
#>

function Set-WatermarkToPDF
{
  <#
      .Synopsis
      Set a watermark to a PDF
      .DESCRIPTION
      You can set a given watermark from an image file to a PDF.
      Output is a new PDF with a watermark.
      .EXAMPLE
      Set-WatermarkToPDF -Name My.pdf -Output My_Copy.pdf -Watermark watermark.png -SetAbsolutePositionXY 0,600 # 100,300

  #>
  Param
  (
    # Insert filename
    [String]
    [Parameter(
        Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
      Position=0)
    ]
    $Name,
        
    [String]
    [Parameter(
        Mandatory,
      Position=1)
    ]
    $Output,
        
    # File with the watermark
    [String]
    [Parameter(
        Mandatory,
      Position=2)
    ]
    $Watermark,
    
    # Set absolut position of the watermark
    [int[]]
    [Parameter(
        Mandatory,
        Position = 3)
    ]
    $SetAbsolutePositionXY
  )

  Begin
  {
    Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
    $reader  = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $(Join-Path $pwd $Name)
    $memoryStream = New-Object System.IO.MemoryStream
    $pdfStamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, $memoryStream)

    $img = [iTextSharp.text.Image]::GetInstance($Watermark)
    $img.SetAbsolutePosition($SetAbsolutePositionXY[0], $SetAbsolutePositionXY[1])
    [iTextSharp.text.pdf.PdfContentByte]$myWaterMark
  }
  Process
  {    
    $pageIndex = $reader.NumberOfPages
    
    for ($i = 1; $i -le $pageIndex; $i++) {
      $myWaterMark = $pdfStamper.GetOverContent($i)
      $myWaterMark.AddImage($img)
    }
    
    $pdfStamper.FormFlattening = $true
    $pdfStamper.Dispose()

    $bytes = $memoryStream.ToArray()
    $memoryStream.Dispose()
    $reader.Dispose()
    [System.IO.File]::WriteAllBytes($Output, $bytes)
    }
  End {}
}

# Set-WatermarkToPDF demo:
<#
    Set-WatermarkToPDF -Name Christian_Imhorst.pdf -Output Christian_Imhorst_Copy.pdf -Watermark watermark.png -SetAbsolutePositionXY 0,600 # 100,300
    Invoke-Expression $(Join-Path $pwd "Christian_Imhorst_Copy.pdf")

    Set-WatermarkToPDF -Name Christian_Imhorst.pdf -Output Christian_Imhorst_Copy.pdf -Watermark Kopie.png -SetAbsolutePositionXY 100,300
    Invoke-Expression $(Join-Path $pwd "Christian_Imhorst_Copy.pdf")
#>

function ConvertFrom-HtmlToPDF
{
  <#
      .Synopsis
      Convert your HTML page to PDF
      .DESCRIPTION
      Convert your HTML page to PDF with iTextSharp
      .EXAMPLE
      ConvertFrom-HtmlToPDF -HTML sample.html -Output sample.pdf
  #>
    Param
    (
        # Insert your HTML
        [Parameter(Mandatory,
                   ValueFromPipelineByPropertyName,
                   ValueFromPipeline,
                   Position=0)]
        $HTML,

        # PDF file out
        [String]
        [Parameter(Mandatory,
                   Position=1)]
        $Output
    )

    Begin
    {
      Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
      $doc  = New-Object iTextSharp.text.Document
      $memoryStream = New-Object System.IO.MemoryStream
      $null = [itextsharp.text.pdf.PdfWriter]::GetInstance($doc, $memoryStream)
      $example_html = $(Get-Content $(Join-Path $pwd $HTML) | Out-String)
      
    }
    Process
    {
      $doc.Open()

      # Use the built-in HTMLWorker to parse the HTML.
      # Only inline CSS is supported.
      $htmlWorker = New-Object iTextSharp.text.html.simpleparser.HTMLWorker($doc)

      # HTMLWorker doesn't read a string directly but instead needs a TextReader
      $sr = new-object System.IO.StringReader($example_html)
      $htmlWorker.Parse($sr)

      $doc.Close()

      $bytes = $memoryStream.ToArray()
      [System.IO.File]::WriteAllBytes($(Join-Path $pwd $Output), $bytes)
    }
    End {}
}

# ConvertFrom-HtmlToPDF demo: 
<#
    ConvertFrom-HtmlToPDF -HTML sample.html -Output sample.pdf
    Invoke-Expression $(Join-Path $pwd "sample.pdf")
#>

function Join-PDFFiles
{
  <#
      .Synopsis
      Join PDF files to one PDF
      .DESCRIPTION
      Join PDF files to one PDF with iTextSharp
      .EXAMPLE
      Join-PDFFiles -Filenames $(gci *.pdf) -Output "JoinedPDFs.pdf"
  #>
  param
  (
    [string[]]
    [Parameter(Mandatory,
        ValueFromPipeline,
        ValueFromPipelineByPropertyName,
      Position=0)
    ]
    $Filenames,

    [String]
    [Parameter(Mandatory,
    Position=1)]
    $Output
  )

  begin
  {
    Add-Type -Path $(Join-Path $pwd "itextsharp.dll")
    $doc  = New-Object iTextSharp.text.Document
    $fs = [System.IO.FileStream]::new($(Join-Path $pwd $Output), [System.IO.FileMode]::Create)
    $writer = New-Object iTextSharp.text.pdf.PdfCopy($doc, $fs)
    $doc.Open()
  }
  process
  {
    foreach ($filename in $filenames)
    {
      $reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $filename
      $reader.ConsolidateNamedDestinations()
  
      for ($i = 1; $i -le $reader.NumberOfPages; $i++) 
      {
        $page = $writer.GetImportedPage($reader, $i)
        $writer.AddPage($page)
      }
      $reader.Close()
  
    }
  }
  end
  {
    $writer.Close()
    $doc.Close()
  }
}

# Join-PDFFiles: 
<#
    Join-PDFFiles -Filenames $(gci *.pdf) -Output JoinedPDFs.pdf
    Invoke-Expression $(Join-Path $pwd "JoinedPDFs.pdf")
#>
