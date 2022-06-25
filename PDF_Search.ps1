
# PDP_Search
# A powershell script to search a file list in an excel file for a set of strings listed in another file.  The found output is a .csv file
#
# Author T. Robinson
# 6/18/2022
#
# 1. Read the string table to get a list of strings to search
# 2. Search each file listed in the source file for presence of any search string
# 3. Output matches as a CSV to the output file
#
# args[0] - the search string file (full path)
# args[1] - the search string sheet name 
# args[2] - the source file (This list of files - all should be full path)
# args[3] - the source file sheet
# args[4] - the output file name (full path)

#Get the Search Table and the sheet for the Search
# the Search is in column 1
$file_name = $args[0]
$sheet_name = $args[1]

Write-Host  "file_name = $file_name"
Write-Host  "sheet_name = $sheet_name"

# open and get the Search table worksheet
$objExcel = New-Object -ComObject Excel.Application  
$WorkBook = $objExcel.Workbooks.Open($file_name,$null,$true)  
$WorkBook.sheets | Select-Object -Property Name  
$WorkSheet = $WorkBook.sheets.Item($sheet_name)

$totalNoOfRecords = $Worksheet.usedRange.Rows.count

Write-Host "looping through $file_name with  $totalNoOfRecords Search Records"

# open the file with the file list to 
$source_file = $args[2]
Write-Host "source file = $source_file"

$file_sheet = $args[3]
Write-Host "file sheet = $file_sheet" 

$FileBook = $objExcel.Workbooks.Open($source_File,$null,$true)  
$FileBook.sheets | Select-Object -Property Name  
$FileSheet = $FileBook.sheets.Item($file_sheet)
$totalNoOfFiles = $FileSheet.usedRange.Rows.count

$out_file = $args[4]
Write-Host "out_file  = $out_file"
$Stream = [System.IO.StreamWriter]::new($out_file)
Write-Host("writing to $out_file")

Write-Host "searching file $source_file with $totalNoOfFiles files"

# change this to where iText Sharp is stored
$iTextSharpFilePath = "C:\temp\itextsharp\itextsharp.dll"
$iBouncyCastleFilePath = "C:\temp\itextsharp\BouncyCastle.Crypto.dll"

[System.Reflection.Assembly]::LoadFrom($iBouncyCastleFilePath)
[System.Reflection.Assembly]::LoadFrom($iTextSharpFilePath)
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# loop through all the files
for ( $i = 2 ; $i -le $totalNoOfFiles; $i++ ) 
{
    # extract path and the part - path is column 4 and part is column 1
    $file = $FileSheet.Columns.Item(4).Rows.Item($i).text
    $part = $FileSheet.Columns.Item(1).Rows.Item($i).text

    try {
        $pdf = [iTextSharp.text.pdf.PdfReader]::new($file)
        Write-Host (" Processing file Name $file" )
        $found_count = 0
    } catch {
        Write-Host (" File Error file Name $file" )
        $Stream.WriteLine( ";;$file;Error Opening ")
    }
    $pdf = [iTextSharp.text.pdf.PdfReader]::new($file)
    Write-Host (" Processing file Name $file" )
    $found_count = 0

    foreach ($Page in 1..($pdf.NumberOfPages))
    {
        $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$Page)
        for ( $j = 2 ; $j -le $totalNoOfRecords; $j++ ) 
        {   
            $string = $WorkSheet.Columns.Item(1).Rows.Item($j).text
            #Write-Host ("Search String $string")
            try{
                $pageText | Select-String -Pattern $string | 
                ForEach-Object { 
                    # Write-Host "$string;$file" 
                    $Stream.WriteLine( "$part;$string;$file")
                    $found_count++
                }
            } catch {
                Write-Host "Error occurred for $file with $string"
                $Stream.WriteLine( "$part;$string;$file;Error Occurred")
            }          
        }
    }
    if($found_count -eq 0){
        Write-Host "$file No Searchs Found" 
        $Stream.WriteLine( "$part;No Searchs;$file")
    }
    Write-Host "Processed $file with $found_count Search references"
}
  

$Stream.close()