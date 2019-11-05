$projectRoot = (Get-Item $PSScriptRoot).Parent
Set-Location $projectRoot.FullName
git tag | foreach-object -process { 
    git push --delete origin $_
    git tag -d $_ 
}