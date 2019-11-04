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
$latestVersion = $existing | Sort-Object LastAccessTime -Descending | Select-Object -First 1
if ($latestVersion) {
    Write-Host $latest.name
    $currentVersion = [regex]::Match($latestVersion.Name,"(\d.\d.\d)").captures.groups[1].value
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
$currentVersion = $versionArray -join "."

$standaloneList = $standaloneList.ForEach({"$src\$_"})
$withTestsList = $withTestsList.ForEach({"$src\$_"})

$outputPath = New-Item -ItemType Directory -Force -Path (Join-Path -Path $releases.FullName -ChildPath "v$($currentVersion)")

$standalonePath = "$($outputPath.FullName)\Standalone.Zip"
$withTestsPath  = "$($outputPath.FullName)\With Tests.Zip"

# Create .zip files
Compress-Archive -Path $standaloneList -CompressionLevel Optimal -DestinationPath $standalonePath
Compress-Archive -Path $withTestsList -CompressionLevel Optimal -DestinationPath $withTestsPath

# Delete current files in latest
Get-ChildItem -Path $latest.FullName | Remove-Item -Recurse

# Copy new files to latest
Copy-Item -Path $standalonePath -Destination $latest.FullName
Copy-Item -Path $withTestsPath -Destination $latest.FullName

Set-Location $projectRoot.FullName
git add --all
git commit --message $currentVersion
git tag $currentVersion
git push


return $currentVersion
