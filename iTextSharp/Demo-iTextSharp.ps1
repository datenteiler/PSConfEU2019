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

function Set-WatermarkToPDF
{
  <#
      .Synopsis
      Set a watermark to a PDF
      .DESCRIPTION
      You can set a given watermark from an image file to a PDF.
      Output is a new PDF with a watermark.
      .EXAMPLE
      Set-WatermarkToPDF -File My.pdf -Output My_Copy.pdf -Watermark watermark.png -SetAbsolutePositionXY 0,600 # 100,300

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
  }
  Process
  {    
    $memoryStream = New-Object System.IO.MemoryStream
    $pdfStamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, $memoryStream)

    $img = [iTextSharp.text.Image]::GetInstance($Watermark)
    $img.SetAbsolutePosition($SetAbsolutePositionXY[0], $SetAbsolutePositionXY[1])
    [iTextSharp.text.pdf.PdfContentByte]$myWaterMark

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
