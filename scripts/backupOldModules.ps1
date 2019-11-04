$projectRoot = (Get-Item $PSScriptRoot).Parent
$src = Get-Item (Join-Path -Path $projectRoot.FullName -ChildPath "src")
$currentItems = Get-ChildItem -Path $src
if($currentItems) {
    $old = Get-Item (Join-Path -Path $projectRoot.FullName -ChildPath "old")
    $oldDate = $src.LastWriteTime.ToString("yyyy-MM-dd")
    $destination = Join-Path -Path $old.FullName -ChildPath $oldDate

    $i = 0
    While (Test-Path $destination) {
        $i += 1
        $destination = Join-Path -Path $old.FullName -ChildPath "$($oldDate) ($($i))"
    }
    New-Item -ItemType Directory -Path $destination
    $currentItems | Move-Item -Destination $destination
}


