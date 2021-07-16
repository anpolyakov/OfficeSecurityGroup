# Создание группы для Microsoft Office
New-LocalGroup -Name "MSOffice_Users" -ErrorAction SilentlyContinue

# Добавляет перенос строки
$OFS = "`n"

# Создается массив
[string[]] $_UserList= @()

# Получение списка пользователей от пользователя
$_UserList = Read-Host "Введите через запятую список пользователей для Microsoft Office"

# Разделение строк из строки в список
$_UserList = $_UserList.Split(',').Split('') > $env:temp\users.txt

# Удаление пробелов и переносов из файла
(gc $env:temp\users.txt) | ? {$_.trim() -ne "" } | set-content $env:temp\users.txt

# Создание массива из файла
[string[]]$users = @(Get-Content -Path $env:temp\users.txt)

# Перебор массива и добавление пользователей в группу
foreach ($user in $users) {
  Add-LocalGroupMember -Group "MSOffice_Users" -Member $user
}

# Массив путей к офисному пакету
$pathes = @(
    "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
    "C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
    "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
    "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"
)

# Цикл назначения прав
foreach ($path in $pathes) {
  if (Test-Path -Path $path)
  {
    $acl = Get-Acl $path -ErrorAction SilentlyContinue
    $acl.SetAccessRuleProtection($true,$true) 
    $acl | Set-Acl $path
    $acl = Get-Acl $path -ErrorAction SilentlyContinue
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Пользователи","ReadAndExecute","Allow")
    $acl.RemoveAccessRule($AccessRule)
    $acl | Set-Acl $path
    $acl = Get-Acl $path -ErrorAction SilentlyContinue
    $MSOfficeAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("MSOffice_Users", "Modify", "Allow")
    $acl.AddAccessRule($MSOfficeAccessRule)
    $acl | Set-Acl $path
  }
}
