.\env.ps1

# バッチ実行関数
function Invoke-Batch {
    param (
        [string]$target_path
    )
    
    Set-Location -Path $PSScriptRoot
    Write-Host "$target_path Start"
    py -m search_docs $target_path
    Write-Host "$target_path Completed"
}

# # redmine_exportの実行
# Invoke-Batch -target_path ./test
