$projectRoot = (Get-Item $PSScriptRoot).Parent
Set-Location $projectRoot.FullName
git tag | foreach-object -process { 
    git push --delete origin tagName
    git tag -d $_ 
}