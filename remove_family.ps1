# 현재 폴더 내의 모든 .xlsx 파일을 대상으로 작업 수행
$xlsxFiles = Get-ChildItem -Path (Get-Location) -Filter *.xlsx

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

    # 압축하여 원본 파일 대체
    $updatedXlsxPath = Join-Path -Path (Get-Location) -ChildPath "$($xlsxFile.BaseName)-updated.xlsx"
    if (Test-Path $updatedXlsxPath) { Remove-Item -Path $updatedXlsxPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($unzipFolder, $updatedXlsxPath)

    Write-Output "Updated .xlsx file has been created at: $updatedXlsxPath"

    # 임시 폴더 삭제
    Remove-Item -Recurse -Force -Path $unzipFolder
}
