<#
Name         : Waardepapieren Importfiles Conversie Script
Author       : Andreas Reinicke
Organisation : Dienst IT, Fontys Hogescholen
Version      : 1.0
Date         : 03-05-2019
#>

#region Functions

Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

#endregion Functions

# Kies de invoerfile en lees deze in
$InFileName = Get-FileName -initialDirectory "c:fso"
$lijst = Import-Csv $InFileName -Delimiter ";"

# Bepaal de naam van de uitvoerfile(s)
$OutFileNameBase = Read-Host "Geef de basis-filenaam van de te genereren files op"
if ($OutFileNameBase.Length -eq 0)
{
   $OutFileNameBase = "Uitvoer"
}
$OutFileNameExt = ".txt"

# Initialisates
$MaxItems = 300      # Maximaal aantal regels per file
$FileCounter = 1     # Teller die het nummer van de uitvoerfile bijhoud
$ItemCount = 0       # Teller die bijhoud hoeveel regels al in de huidige uitvoerfile staan
$LineCount = 0       # Teller die bijhoud hoeveel regels in totaal reeds verwerkt zijn
$ScriptPath = Split-path $MyInvocation.MyCommand.Definition
$Koptekst = Get-Content $Scriptpath\OutHeader.txt

# Hoofdprogramma
foreach ($item in $lijst)
{
   If ($ItemCount -eq 0)
   {
      $Uitvoer = Get-Content $Scriptpath\OutPrefix.txt
   }
   $Uitvoer += $item.Documentnummer+";"+$item.Relatie[0]+$item.Relatie[1]+"-"+$item.Documentnummer
   $ItemCount += 1
   $LineCount += 1
   If ($LineCount -lt $lijst.Count -and $ItemCount -ne $MaxItems)
   {
      $Uitvoer += "|"
   }
   If ($ItemCount -eq $MaxItems -or $LineCount -eq $lijst.Count)
   {
      $OutFileName = $OutFileNameBase+"-"+$FileCounter.ToString("000")+$OutFileNameExt
      $Koptekst | Set-Content $OutFileName -Encoding Unicode
      $Uitvoer | Add-Content $OutFileName -NoNewline -Encoding Unicode
      If ($LineCount -ne $lijst.Count)
      {
         $ItemCount = 0
         $FileCounter += 1
      }
   }
}