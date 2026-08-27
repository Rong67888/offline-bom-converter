[CmdletBinding()]
param(
    [switch]$LocalCheckOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepoRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$ResultFile = Join-Path $RepoRoot 'github_upload_result.txt'
$ExpectedOwner = 'Rong67888'
$ExpectedName = 'offline-bom-converter'
$ExpectedRepository = "$ExpectedOwner/$ExpectedName"
$ExpectedHttpsRemote = "https://github.com/$ExpectedRepository.git"
$ResultLines = New-Object 'System.Collections.Generic.List[string]'
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$OriginalPath = $env:PATH
$OriginalPythonPath = $env:PYTHONPATH
$OriginalNoBytecode = $env:PYTHONDONTWRITEBYTECODE

function Add-ResultLine {
    param([string]$Text)
    [void]$script:ResultLines.Add($Text)
}

function Save-ResultFile {
    [System.IO.File]::WriteAllLines($script:ResultFile, $script:ResultLines, $script:Utf8Bom)
}

function Find-Executable {
    param(
        [string]$CommandName,
        [string[]]$CandidatePaths
    )
    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($null -ne $command -and $command.Source) {
        return $command.Source
    }
    foreach ($candidate in $CandidatePaths) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }
    return $null
}

function Convert-ToSafeErrorText {
    param([object[]]$Output)
    $text = (($Output | ForEach-Object { $_.ToString() }) -join "`n").Trim()
    $text = [regex]::Replace($text, 'gh[pousr]_[A-Za-z0-9]{20,}', '[凭据已隐藏]')
    $text = [regex]::Replace($text, 'github_pat_[A-Za-z0-9_]{20,}', '[凭据已隐藏]')
    $text = [regex]::Replace($text, '(?im)^Authorization:.*$', 'Authorization: [已隐藏]')
    if ($text.Length -gt 800) {
        $text = $text.Substring(0, 800) + '……'
    }
    if ([string]::IsNullOrWhiteSpace($text)) {
        return '命令没有返回可读的错误信息。'
    }
    return $text
}

function Invoke-CapturedCommand {
    param(
        [string]$Executable,
        [string[]]$Arguments
    )
    $output = & $Executable @Arguments 2>&1
    return [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output = @($output)
        Text = (($output | ForEach-Object { $_.ToString() }) -join "`n").Trim()
    }
}

function Assert-ExpectedPrivateRepository {
    param([object]$RepositoryData)
    $actualOwner = [string]$RepositoryData.owner.login
    $actualName = [string]$RepositoryData.name
    $actualVisibility = ([string]$RepositoryData.visibility).ToUpperInvariant()
    if (-not $actualOwner.Equals($script:ExpectedOwner, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "仓库所有者不符合预期：应为 $($script:ExpectedOwner)，实际为 $actualOwner。"
    }
    if (-not $actualName.Equals($script:ExpectedName, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "仓库名称不符合预期：应为 $($script:ExpectedName)，实际为 $actualName。"
    }
    if ($RepositoryData.'private' -ne $true -or $actualVisibility -ne 'PRIVATE') {
        throw '目标仓库不是Private。脚本已停止，不会推送，也不会修改仓库可见性。'
    }
}

try {
    Set-Location -LiteralPath $RepoRoot
    Add-ResultLine 'GitHub私有仓库上传结果'
    Add-ResultLine ("执行时间：{0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'))
    Add-ResultLine "目标仓库：$ExpectedRepository"
    Add-ResultLine '要求可见性：PRIVATE'

    if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot '.git') -PathType Container)) {
        throw '当前目录不是准备好的Git仓库，已停止。'
    }

    $gitCandidates = @(
        (Join-Path $env:ProgramFiles 'Git\cmd\git.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Git\cmd\git.exe')
    )
    $GitExe = Find-Executable -CommandName 'git.exe' -CandidatePaths $gitCandidates
    if (-not $GitExe) {
        throw '没有找到Git。请先安装Git for Windows，然后重新双击BAT。'
    }

    $pythonCandidates = @(
        (Join-Path $RepoRoot '.venv\Scripts\python.exe'),
        (Join-Path $RepoRoot '.venv-build\Scripts\python.exe')
    )
    $PythonExe = Find-Executable -CommandName 'python.exe' -CandidatePaths $pythonCandidates
    if (-not $PythonExe) {
        throw '没有找到Python。上传前必须用Python重新运行测试和脱敏审计。'
    }

    $gitFolder = Split-Path -Parent $GitExe
    $env:PATH = "$gitFolder;$OriginalPath"
    $env:PYTHONPATH = Join-Path $RepoRoot 'src'
    $env:PYTHONDONTWRITEBYTECODE = '1'

    $branchResult = Invoke-CapturedCommand -Executable $GitExe -Arguments @('-C', $RepoRoot, 'branch', '--show-current')
    if ($branchResult.ExitCode -ne 0 -or $branchResult.Text.Trim() -ne 'main') {
        throw '当前Git分支不是main，已停止。'
    }

    $headResult = Invoke-CapturedCommand -Executable $GitExe -Arguments @('-C', $RepoRoot, 'rev-parse', 'HEAD')
    if ($headResult.ExitCode -ne 0 -or $headResult.Text.Trim() -notmatch '^[0-9a-fA-F]{40}$') {
        throw '没有找到有效的本地Git提交，已停止。'
    }
    $LocalSha = $headResult.Text.Trim().ToLowerInvariant()

    $statusResult = Invoke-CapturedCommand -Executable $GitExe -Arguments @('-C', $RepoRoot, 'status', '--porcelain=v1', '--untracked-files=all')
    if ($statusResult.ExitCode -ne 0) {
        throw '无法检查Git工作区状态，已停止。'
    }
    if (-not [string]::IsNullOrWhiteSpace($statusResult.Text)) {
        throw 'Git工作区存在未提交或未跟踪的文件。请不要上传，先检查并提交或移走这些文件。'
    }

    Write-Host '步骤1/3：运行全部公开版自动化测试……' -ForegroundColor Cyan
    & $PythonExe -m unittest discover -s (Join-Path $RepoRoot 'tests') -v
    if ($LASTEXITCODE -ne 0) {
        throw '自动化测试未全部通过，已停止上传。'
    }
    Add-ResultLine '自动化测试：通过'

    Write-Host '步骤2/3：运行Git跟踪文件、隐私和XLSX元数据审计……' -ForegroundColor Cyan
    & $PythonExe (Join-Path $RepoRoot 'tools\audit_public_repo.py') --require-git-tracked
    if ($LASTEXITCODE -ne 0) {
        throw '脱敏审计未通过，已停止上传。'
    }
    Add-ResultLine '脱敏与Git跟踪审计：通过'

    $statusAfterChecks = Invoke-CapturedCommand -Executable $GitExe -Arguments @('-C', $RepoRoot, 'status', '--porcelain=v1', '--untracked-files=all')
    if ($statusAfterChecks.ExitCode -ne 0 -or -not [string]::IsNullOrWhiteSpace($statusAfterChecks.Text)) {
        throw '测试后Git工作区不再干净，已停止上传。'
    }
    Add-ResultLine "本地main提交：$LocalSha"

    if ($LocalCheckOnly) {
        Add-ResultLine '状态：本地检查通过；未连接GitHub，未创建仓库，未推送。'
        Save-ResultFile
        Write-Host '本地检查通过。本次没有执行任何GitHub或互联网操作。' -ForegroundColor Green
        exit 0
    }

    Write-Host '步骤3/3：检查GitHub登录并安全推送Private仓库……' -ForegroundColor Cyan
    $ghCandidates = @(
        (Join-Path $env:ProgramFiles 'GitHub CLI\gh.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\GitHub CLI\gh.exe')
    )
    $GhExe = Find-Executable -CommandName 'gh.exe' -CandidatePaths $ghCandidates
    if (-not $GhExe) {
        throw '没有找到GitHub CLI。请先安装并登录GitHub CLI，然后重新运行。'
    }

    & $GhExe auth status --hostname github.com *> $null
    if ($LASTEXITCODE -ne 0) {
        throw 'GitHub CLI尚未登录或登录已失效。请在本机完成登录，不要在聊天或文件中填写Token。'
    }

    $loginResult = Invoke-CapturedCommand -Executable $GhExe -Arguments @('api', 'user', '--jq', '.login')
    if ($loginResult.ExitCode -ne 0) {
        throw ('无法核对GitHub登录账号：' + (Convert-ToSafeErrorText $loginResult.Output))
    }
    $actualLogin = $loginResult.Text.Trim()
    if (-not $actualLogin.Equals($ExpectedOwner, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "当前GitHub登录账号为 $actualLogin，不是预期账号 $ExpectedOwner，已停止。"
    }
    Add-ResultLine "GitHub账号：$actualLogin"

    $repoResult = Invoke-CapturedCommand -Executable $GhExe -Arguments @('api', "repos/$ExpectedRepository")
    if ($repoResult.ExitCode -eq 0) {
        $repoData = $repoResult.Text | ConvertFrom-Json
        Assert-ExpectedPrivateRepository -RepositoryData $repoData
        Write-Host '目标Private仓库已经存在，已核对所有者、名称和可见性。' -ForegroundColor Yellow
    }
    else {
        $safeRepoError = Convert-ToSafeErrorText $repoResult.Output
        if ($safeRepoError -notmatch '(?i)(HTTP\s+404|Not Found)') {
            throw "无法安全判断目标仓库是否存在：$safeRepoError"
        }
        Write-Host '目标仓库不存在，正在创建Private仓库……' -ForegroundColor Yellow
        $createResult = Invoke-CapturedCommand -Executable $GhExe -Arguments @(
            'repo', 'create', $ExpectedRepository, '--private',
            '--description', 'Privacy-safe preview of an offline XLSX BOM converter'
        )
        if ($createResult.ExitCode -ne 0) {
            throw ('创建Private仓库失败：' + (Convert-ToSafeErrorText $createResult.Output))
        }
        $repoResult = Invoke-CapturedCommand -Executable $GhExe -Arguments @('api', "repos/$ExpectedRepository")
        if ($repoResult.ExitCode -ne 0) {
            throw ('仓库创建后无法重新核对：' + (Convert-ToSafeErrorText $repoResult.Output))
        }
        $repoData = $repoResult.Text | ConvertFrom-Json
        Assert-ExpectedPrivateRepository -RepositoryData $repoData
    }

    $existingDefaultBranch = [string]$repoData.default_branch
    if ($existingDefaultBranch -and $existingDefaultBranch -ne 'main') {
        throw "目标仓库的默认分支不是main（当前为 $existingDefaultBranch）。为避免推送到意外仓库，脚本已停止。"
    }

    $originResult = Invoke-CapturedCommand -Executable $GitExe -Arguments @('-C', $RepoRoot, 'remote', 'get-url', 'origin')
    if ($originResult.ExitCode -ne 0) {
        & $GitExe -C $RepoRoot remote add origin $ExpectedHttpsRemote
        if ($LASTEXITCODE -ne 0) {
            throw '无法添加安全的origin远端地址，已停止。'
        }
    }
    else {
        $originUrl = $originResult.Text.Trim()
        $allowedOrigins = @(
            $ExpectedHttpsRemote,
            "https://github.com/$ExpectedRepository",
            "git@github.com:$ExpectedRepository.git"
        )
        if ($allowedOrigins -notcontains $originUrl) {
            throw "现有origin指向其他地址：$originUrl。为避免推错仓库，脚本不会自动修改。"
        }
    }

    & $GitExe -C $RepoRoot push --set-upstream origin main
    if ($LASTEXITCODE -ne 0) {
        throw 'Git推送失败。脚本没有强制推送，也没有改写远端历史；请查看上方Git提示后重试。'
    }

    $finalRepoResult = Invoke-CapturedCommand -Executable $GhExe -Arguments @('api', "repos/$ExpectedRepository")
    if ($finalRepoResult.ExitCode -ne 0) {
        throw ('推送后无法核对仓库信息：' + (Convert-ToSafeErrorText $finalRepoResult.Output))
    }
    $finalRepo = $finalRepoResult.Text | ConvertFrom-Json
    Assert-ExpectedPrivateRepository -RepositoryData $finalRepo

    $remoteCommitResult = Invoke-CapturedCommand -Executable $GhExe -Arguments @('api', "repos/$ExpectedRepository/commits/main", '--jq', '.sha')
    if ($remoteCommitResult.ExitCode -ne 0) {
        throw ('无法读取远端main最新提交：' + (Convert-ToSafeErrorText $remoteCommitResult.Output))
    }
    $RemoteSha = $remoteCommitResult.Text.Trim().ToLowerInvariant()
    if ($RemoteSha -ne $LocalSha) {
        throw "远端main提交与本地不一致。本地：$LocalSha；远端：$RemoteSha。"
    }

    $defaultBranch = [string]$finalRepo.default_branch
    if ($defaultBranch -ne 'main') {
        throw "仓库已经上传，但默认分支不是main（当前为 $defaultBranch）。脚本不会自动修改仓库设置。"
    }

    $repoUrl = [string]$finalRepo.html_url
    Add-ResultLine "仓库URL：$repoUrl"
    Add-ResultLine '可见性：PRIVATE'
    Add-ResultLine "默认分支：$defaultBranch"
    Add-ResultLine "远端main提交：$RemoteSha"
    Add-ResultLine '本地与远端提交：一致'
    Add-ResultLine '状态：上传成功，仓库仍为Private'
    Save-ResultFile

    Write-Host ''
    Write-Host '上传成功，仓库仍为Private' -ForegroundColor Green
    Write-Host "仓库地址：$repoUrl"
    Write-Host "最新提交：$RemoteSha"
    exit 0
}
catch {
    $reason = Convert-ToSafeErrorText @($_.Exception.Message)
    Add-ResultLine "状态：上传失败"
    Add-ResultLine "失败原因：$reason"
    try {
        Save-ResultFile
    }
    catch {
        Write-Host '同时无法写入github_upload_result.txt，请检查目录写入权限。' -ForegroundColor Red
    }
    Write-Host ''
    Write-Host "上传失败：$reason" -ForegroundColor Red
    exit 1
}
finally {
    $env:PATH = $OriginalPath
    $env:PYTHONPATH = $OriginalPythonPath
    $env:PYTHONDONTWRITEBYTECODE = $OriginalNoBytecode
}
