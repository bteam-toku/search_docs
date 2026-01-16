# 一つ上の親ディレクトリをカレントフォルダリに設定（環境に合わせて変更してください）
Set-Location -Path (Join-Path $PSScriptRoot "..")
# venvの有効化
.\scripts\env.ps1

# バッチ実行関数
function Invoke-Batch {
    param (
        [string]$target_path
    )
    py -m search_docs $target_path
}

# redmine_exportの実行
Invoke-Batch -target_path ./test
