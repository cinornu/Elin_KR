# 현재 폴더 내의 모든 .xlsx 파일을 대상으로 작업 수행
$xlsxFiles = Get-ChildItem -Path (Get-Location) -Filter *.xlsx
$updatedFiles = @() # -updated 파일 목록을 저장할 배열

foreach ($xlsxFile in $xlsxFiles) {
    Write-Output "Processing file: $($xlsxFile.FullName)"

    # 임시 폴더 경로 설정
    $unzipFolder = Join-Path -Path (Get-Location) -ChildPath "temp_unzip"
    if (Test-Path -Path $unzipFolder) { Remove-Item -Recurse -Force -Path $unzipFolder }
    New-Item -ItemType Directory -Path $unzipFolder | Out-Null

    # 압축 해제
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($xlsxFile.FullName, $unzipFolder)

    # 텍스트 기반으로 <family/> 태그 제거 함수
    function Remove-FamilyTag {
        param (
            [string]$xmlPath
        )

        if (Test-Path $xmlPath) {
            # XML 파일을 텍스트로 로드하여 <family/> 태그 삭제
            $xmlContent = Get-Content -Path $xmlPath -Raw -Encoding UTF8
            $xmlContent = $xmlContent -replace "<family\s*/>", ""
            Set-Content -Path $xmlPath -Value $xmlContent -Encoding UTF8
            Write-Output "Removed <family/> tags from $xmlPath"
        }
    }

    # styles.xml과 sharedStrings.xml에서 <family/> 태그 제거
    $stylesXmlPath = Join-Path -Path $unzipFolder -ChildPath "xl\styles.xml"
    $sharedStringsXmlPath = Join-Path -Path $unzipFolder -ChildPath "xl\sharedStrings.xml"

    Remove-FamilyTag -xmlPath $stylesXmlPath
    Remove-FamilyTag -xmlPath $sharedStringsXmlPath

    # 압축하여 -updated 파일 생성
    $updatedXlsxPath = Join-Path -Path (Get-Location) -ChildPath "$($xlsxFile.BaseName)-updated.xlsx"
    if (Test-Path $updatedXlsxPath) { Remove-Item -Path $updatedXlsxPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($unzipFolder, $updatedXlsxPath)

    Write-Output "Updated .xlsx file has been created at: $updatedXlsxPath"

    # 임시 폴더 삭제
    Remove-Item -Recurse -Force -Path $unzipFolder

    # -updated 파일 목록에 추가
    $updatedFiles += [PSCustomObject]@{Original = $xlsxFile.FullName; Updated = $updatedXlsxPath}
}

# 패치 적용 여부 확인 메시지
$userInput = Read-Host "Do you want to replace all original files with the updated versions? (Y/N)"
if ($userInput -eq "Y") {
    foreach ($file in $updatedFiles) {
        # 원본 파일 삭제 및 -updated 파일 이름 변경
        if (Test-Path -Path $file.Original) {
            Remove-Item -Path $file.Original -Force
            Rename-Item -Path $file.Updated -NewName (Split-Path -Path $file.Original -Leaf)
            Write-Output "Replaced original file with the updated version: $file.Original"
        }
    }
} else {
    Write-Output "Skipped replacing original files. Updated files remain in their -updated versions."
}
