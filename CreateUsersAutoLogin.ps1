Import-Module ActiveDirectory
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ==============================
# НАСТРОЙКИ
# ==============================
$Domain  = "home.lab"
$BaseOU  = "OU=Company,DC=home,DC=lab"
$CsvPath = "C:\Scripts\users.csv"
# ⚠ Demo password (change in production!)
$DefaultPassword = ConvertTo-SecureString "Start123!" -AsPlainText -Force


# ==============================
# ТРАНСЛИТЕРАЦИЯ (100% рабочая)
# ==============================
function ConvertTo-Latin {
    param([string]$Text)

    $Text = $Text.ToLower()

    $map = @{
        'а'='a'; 'б'='b'; 'в'='v'; 'г'='g'; 'д'='d'
        'е'='e'; 'ё'='e'; 'ж'='zh'; 'з'='z'; 'и'='i'
        'й'='y'; 'к'='k'; 'л'='l'; 'м'='m'; 'н'='n'
        'о'='o'; 'п'='p'; 'р'='r'; 'с'='s'; 'т'='t'
        'у'='u'; 'ф'='f'; 'х'='kh'; 'ц'='ts'; 'ч'='ch'
        'ш'='sh'; 'щ'='sch'; 'ы'='y'; 'э'='e'; 'ю'='yu'
        'я'='ya'; 'ь'=''; 'ъ'=''
    }

    $result = ""

     foreach ($c in $Text.ToCharArray()) {

        $ch = $c.ToString()

        if ($map.ContainsKey($ch)) {
            $result += $map[$ch]
        }
        else {
            $result += $ch
        }
    }

    return $result
}

# ==============================
# ТЕСТ ТРАНСЛИТА (важно!)
# ==============================
Write-Host "`n=== TEST TRANSLIT ==="
Write-Host "Иванов -> $(ConvertTo-Latin "Иванов")"
Write-Host "Петров -> $(ConvertTo-Latin "Петров")"
Write-Host "=====================`n"


# ==============================
# ИМПОРТ CSV
# ==============================
if (!(Test-Path $CsvPath)) {
    Write-Host "❌ CSV файл не найден: $CsvPath"
    exit
}

$Users = Import-Csv $CsvPath -Encoding UTF8


# ==============================
# СОЗДАНИЕ ПОЛЬЗОВАТЕЛЕЙ
# ==============================
foreach ($User in $Users) {

    $FirstName = $User.FirstName.Trim()
    $LastName  = $User.LastName.Trim()
    $Dept      = $User.Department.Trim()

    $FullName = "$FirstName $LastName"

    # Генерация логина: фамилия + первая буква имени
    $Login = (ConvertTo-Latin $LastName) + "." + (ConvertTo-Latin $FirstName.Substring(0,1))

    # Проверка: логин должен быть латиницей
    if ($Login -match "[а-яё]") {
        Write-Host "❌ ОШИБКА: логин содержит русские буквы: $Login"
        continue
    }

    $TargetOU = "OU=$Dept,$BaseOU"

    # Проверка существования OU
    if (!(Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$TargetOU'" -ErrorAction SilentlyContinue)) {
        Write-Host "⚠ OU не найдена: $TargetOU — пропуск пользователя $FullName"
        continue
    }

    # Проверка существования пользователя
    if (Get-ADUser -Filter "SamAccountName -eq '$Login'" -ErrorAction SilentlyContinue) {
        Write-Host "⚠ Пользователь $Login уже существует — пропуск"
        continue
    }

    # Создание пользователя
    New-ADUser `
        -Name $FullName `
        -GivenName $FirstName `
        -Surname $LastName `
        -DisplayName $FullName `
        -SamAccountName $Login `
        -UserPrincipalName "$Login@$Domain" `
        -Path $TargetOU `
        -AccountPassword $DefaultPassword `
        -Enabled $true `
        -ChangePasswordAtLogon $true

    Write-Host "✅ Создан: $FullName → логин: $Login"
}

Write-Host "`n🎉 Импорт завершён!"
