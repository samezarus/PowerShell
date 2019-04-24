# Копирование с перезаписью
#
# 1й параметр - путь откуда копируем 
# 2й параметр - путь куда копируем 
# 3й параметр - маска, по которой копируем 

Clear

#$from = 'D:\qlik\Dev\90_Documentation\'
#$to = 'C:\Program Files\QlikView\Web\help\'
#$mask = '*.html'

$from = $args[0]
$to   = $args[1]
$mask = '*.'+$args[2]

#$from
#$to
#$mask

$filesList = Get-ChildItem -Path $from -Filter $mask

foreach ($item in $filesList)
{
    $s1 = $from+'\'+$item.name
    $s2 = $to+'\'+$item.name
    Copy-Item $s1 -Destination $s2 -Recurse
}
