# IndexNinja
A script leveraging Windows Indexer to find keywords inside files.

## SYNOPSIS
    Gets filenames which have been indexed by Windows desktop search and have contents matching filter keywords
        
    Name: Get-IndexNinja.ps1
    Version: 1.1
    Author: @mgreen27

## DESCRIPTION
Script to search for keywords in Windows Index and output lists for each keyword.

Add in variables in param below - Search path, Results Path and Key Words (text file)

Please ensure all expected items are indexed

Install all appropriate windows filters - in my initial usecase we needed to install a Office 2010 filter to idex xlsx and docx files which are not indexed by default and utilise a binary encoding.

Get-IndexItem from Technet - https://gallery.technet.microsoft.com/scriptcenter/Get-IndexedItem-PowerShell-5bca2dae

Office 2010 Filter update - https://www.microsoft.com/en-us/download/details.aspx?id=17062
    
## EXAMPLE
    Get-IndexNinja.ps1
    Get-IndexNinja.ps1 -SearchPath $SearchPath 
    Get-IndexNinja.ps1 -SearchPath $SearchPath -ResultsPath
    Get-IndexNinja.ps1 -SearchPath $SearchPath -ResultsPath $ResultsPath -list 
