# Скрипт для обработки архивов сделанных Acronis
#
# -pre  - действие перед выполнением бэкапа
#     $backupDir - имя каталога в котором следует удалить все файлы
#
# -post - действие после выполнением бэкапа
#     $backupDir         - имя каталога в котором содержится бэкап (day_2020_12_03_19_00_03_734D.TIB)
#     $newNameBackupFile - новое имя бэкап файла
#
# powershell prePostAcronis.ps1 -pre c:\temp\day\
# powershell prePostAcronis.ps1 -post c:\temp\day\ new.TIB 

clear

function pre_acronis($backupDir)
{
    Remove-Item $backupDir* -Recurse -Force
}

function post_acronis($backupDir, $newNameBackupFile)
{
    $filesList = Get-ChildItem $backupDir
    foreach ($file in $filesList)
    {
        if ($file.Extension -eq ".TIB")
        {
            $s = $backupDir + $file.Name
            Rename-Item -path $s -NewName $newNameBackupFile
        }
    }
}

#post_acronis -backupDir "C:\test\day\" -newNameBackupFile "day.TIB"

if ($args.Count -gt 0)
{
    switch -Exact($args[0])
    {
        '-pre'
        {
            pre_acronis -backupDir $args[1]
        }

        '-post'
        {
            post_acronis -backupDir $args[1] -newNameBackupFile $args[2]
        }
    }
}