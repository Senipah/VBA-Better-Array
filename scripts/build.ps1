param(
    [Parameter(Position=0)]
    [ValidateSet('major','minor','patch')]
    [System.String]$versionIncrement = "patch"
)

$standaloneList = "BetterArray.cls"
$withTestsList = 
    "BetterArray.cls",
    "ArrayGenerator.cls",
    "ExcelProvider.cls",
    "IValuesList.cls",
    "TestModule_ArrayGenerator.bas",
    "TestModule_BetterArray.bas",
    "TestModule_ExcelProvider.bas",
    "ValuesList_Booleans.cls",
    "ValuesList_Bytes.cls", 
    "ValuesList_Doubles.cls", 
    "ValuesList_Longs.cls", 
    "ValuesList_Objects.cls",
    "ValuesList_Strings.cls",
    "ValuesList_Variants.cls"

$projectRoot = (Get-Item $PSScriptRoot).Parent
$src = Get-Item (Join-Path -Path $projectRoot.FullName -ChildPath "src")
$releases = Get-Item (Join-Path -Path $projectRoot.FullName -ChildPath "releases")
$latest= Get-Item (Join-Path -Path $releases.FullName -ChildPath "latest")
$existing =  Get-ChildItem -Path $releases.FullName -Exclude "latest" -Directory 
$previousVersion = $existing | 
    Sort-Object { [version]($_.Name -replace '^.*(\d+(\.\d+){1,3})$', '$1') } -Descending | 
    Select-Object -Index 0
if ($previousVersion) {
    $currentVersion = [regex]::Match($previousVersion.Name,"(\d.\d.\d)").captures.groups[1].value
} else {
    $currentVersion = "0.0.0"
}
$versionArray = $currentVersion.Split(".") 
switch($versionIncrement){
    "major" {
        $versionArray[-1] = 0
        $versionArray[-2] = 0
        $versionArray[-3] = [int]$versionArray[-3] + 1
    }
    "minor" {
        $versionArray[-1] = 0
        $versionArray[-2] = [int]$versionArray[-2] + 1
    }
    "patch" {
        $versionArray[-1] = [int]$versionArray[-1] + 1
    }
}
$currentVersion = "v$($versionArray -join ".")" 
$standaloneList = $standaloneList.ForEach({"$src\$_"})
$nl = [Environment]::NewLine
$previousHeader = "'" + $previousVersion.Name
$currentHeader = "'" + $currentVersion
$withTestsList = $withTestsList.ForEach({
    # Add version number to bottom of all files - standalone is also in this array
    $content = Get-Content "$src\$_"
    if ($content[-1] -ne $currentHeader) {
        if ($content[-1] -eq $previousHeader) {
            $content[-1] = $currentHeader
            $content | Set-Content "$src\$_"
        } else {
            $content + ($nl) + ($currentHeader)  | Set-Content "$src\$_"
        }
    }
    "$src\$_"
})
$outputPath = New-Item -ItemType Directory -Force -Path (Join-Path -Path $releases.FullName -ChildPath $currentVersion)
$standalonePath = "$($outputPath.FullName)\Standalone.Zip"
$withTestsPath  = "$($outputPath.FullName)\WithTests.Zip"

# Create .zip files
Compress-Archive -Path $standaloneList -CompressionLevel Optimal -DestinationPath $standalonePath
Compress-Archive -Path $withTestsList -CompressionLevel Optimal -DestinationPath $withTestsPath

# Delete current files in latest
Get-ChildItem -Path $latest.FullName | Remove-Item -Recurse

# Copy new files to latest
Copy-Item -Path $standalonePath -Destination $latest.FullName
Copy-Item -Path $withTestsPath -Destination $latest.FullName


Set-Location $projectRoot.FullName
$log = git log $previousVersion.Name..HEAD --oneline
Write-Host $log 

git add --all
git commit --message $currentVersion + $nl + $log
git tag $currentVersion
git push
git push --tags
return $currentVersion
Exit