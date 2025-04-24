# Загрузка необходимых сборок для Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Management.Automation

# Проверка наличия модуля Active Directory и его загрузка
if (!(Get-Module -Name ActiveDirectory)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "Модуль ActiveDirectory успешно импортирован"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Невозможно импортировать модуль ActiveDirectory: $($_.Exception.Message)`nПроверьте, установлен ли RSAT (Remote Server Administration Tools).", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Создание синхронизационного хэша для обмена данными между потоками
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.RunspacePool = $null
$syncHash.Jobs = [System.Collections.ArrayList]::new()
$syncHash.SearchResults = $null
$syncHash.SearchCompleted = $false
$syncHash.Error = $null

# Создание основной формы
$form = New-Object System.Windows.Forms.Form
$form.Text = "Поиск групп Active Directory"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"
$form.Icon = [System.Drawing.SystemIcons]::Shield
$syncHash.Form = $form

# Добавление метки для поля поиска
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Location = New-Object System.Drawing.Point(10, 20)
$labelSearch.Size = New-Object System.Drawing.Size(280, 20)
$labelSearch.Text = "Введите шаблон для поиска (например, *доступ*):"
$form.Controls.Add($labelSearch)

# Создание поля ввода для поискового запроса
$textBoxSearch = New-Object System.Windows.Forms.TextBox
$textBoxSearch.Location = New-Object System.Drawing.Point(10, 40)
$textBoxSearch.Size = New-Object System.Drawing.Size(280, 20)
$textBoxSearch.Text = "*доступ*"
$form.Controls.Add($textBoxSearch)
$syncHash.TextBoxSearch = $textBoxSearch

# Создание кнопки поиска
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Location = New-Object System.Drawing.Point(300, 38)
$buttonSearch.Size = New-Object System.Drawing.Size(100, 23)
$buttonSearch.Text = "Найти"
$form.Controls.Add($buttonSearch)
$syncHash.ButtonSearch = $buttonSearch

# Создание группы для выбора метода поиска
$groupBoxSearchMethod = New-Object System.Windows.Forms.GroupBox
$groupBoxSearchMethod.Location = New-Object System.Drawing.Point(410, 10)
$groupBoxSearchMethod.Size = New-Object System.Drawing.Size(360, 60)
$groupBoxSearchMethod.Text = "Метод поиска"
$form.Controls.Add($groupBoxSearchMethod)

# Создание переключателей для методов поиска
$radioStandard = New-Object System.Windows.Forms.RadioButton
$radioStandard.Location = New-Object System.Drawing.Point(10, 20)
$radioStandard.Size = New-Object System.Drawing.Size(100, 30)
$radioStandard.Checked = $true
$radioStandard.Text = "Стандартный"
$groupBoxSearchMethod.Controls.Add($radioStandard)
$syncHash.RadioStandard = $radioStandard

$radioLDAP = New-Object System.Windows.Forms.RadioButton
$radioLDAP.Location = New-Object System.Drawing.Point(120, 20)
$radioLDAP.Size = New-Object System.Drawing.Size(100, 30)
$radioLDAP.Text = "LDAP-фильтр"
$groupBoxSearchMethod.Controls.Add($radioLDAP)
$syncHash.RadioLDAP = $radioLDAP

$radioAdvanced = New-Object System.Windows.Forms.RadioButton
$radioAdvanced.Location = New-Object System.Drawing.Point(230, 20)
$radioAdvanced.Size = New-Object System.Drawing.Size(120, 30)
$radioAdvanced.Text = "Расширенный"
$groupBoxSearchMethod.Controls.Add($radioAdvanced)
$syncHash.RadioAdvanced = $radioAdvanced

# Создание группы для выбора домена
$groupBoxDomain = New-Object System.Windows.Forms.GroupBox
$groupBoxDomain.Location = New-Object System.Drawing.Point(10, 80)
$groupBoxDomain.Size = New-Object System.Drawing.Size(760, 60)
$groupBoxDomain.Text = "Домен для поиска"
$form.Controls.Add($groupBoxDomain)

# Создание выпадающего списка доменов
$comboDomain = New-Object System.Windows.Forms.ComboBox
$comboDomain.Location = New-Object System.Drawing.Point(10, 25)
$comboDomain.Size = New-Object System.Drawing.Size(430, 20)
$comboDomain.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$groupBoxDomain.Controls.Add($comboDomain)
$syncHash.ComboDomain = $comboDomain

# Заполнение выпадающего списка доменов
try {
    $forest = Get-ADForest
    $domains = $forest.Domains
    
    foreach ($domain in $domains) {
        [void]$comboDomain.Items.Add($domain)
    }
    
    # Добавление опции для поиска по всем доменам
    [void]$comboDomain.Items.Add("Все домены")
    
    # Установка текущего домена по умолчанию
    $currentDomain = (Get-ADDomain).DNSRoot
    $comboDomain.SelectedItem = $currentDomain
    if ($comboDomain.SelectedIndex -eq -1 -and $comboDomain.Items.Count -gt 0) {
        $comboDomain.SelectedIndex = 0
    }
    $syncHash.Domains = $domains
} catch {
    [void]$comboDomain.Items.Add("Ошибка получения доменов")
    $comboDomain.SelectedIndex = 0
    Write-Host "Ошибка при получении списка доменов: $($_.Exception.Message)"
}

# Кнопка для получения инфо о домене
$buttonDomainInfo = New-Object System.Windows.Forms.Button
$buttonDomainInfo.Location = New-Object System.Drawing.Point(450, 25)
$buttonDomainInfo.Size = New-Object System.Drawing.Size(130, 23)
$buttonDomainInfo.Text = "Инфо о домене"
$groupBoxDomain.Controls.Add($buttonDomainInfo)

# Кнопка проверки подключения
$buttonTestConnection = New-Object System.Windows.Forms.Button
$buttonTestConnection.Location = New-Object System.Drawing.Point(590, 25)
$buttonTestConnection.Size = New-Object System.Drawing.Size(150, 23)
$buttonTestConnection.Text = "Проверить подключение"
$groupBoxDomain.Controls.Add($buttonTestConnection)

# Создание таблицы результатов (DataGridView)
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 150)
$dataGridView.Size = New-Object System.Drawing.Size(760, 400)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.ReadOnly = $true
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.AllowUserToOrderColumns = $true
$dataGridView.MultiSelect = $false
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$form.Controls.Add($dataGridView)
$syncHash.DataGridView = $dataGridView

# Добавление журнала (лога)
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 560)
$logTextBox.Size = New-Object System.Drawing.Size(590, 40)
$logTextBox.Multiline = $true
$logTextBox.ReadOnly = $true
$logTextBox.ScrollBars = "Vertical"
$form.Controls.Add($logTextBox)
$syncHash.LogTextBox = $logTextBox

# Индикатор прогресса
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 150)
$progressBar.Size = New-Object System.Drawing.Size(760, 20)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 0
$progressBar.Visible = $false
$form.Controls.Add($progressBar)
$syncHash.ProgressBar = $progressBar

# Добавление статусной строки
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Готов к поиску"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)
$syncHash.StatusLabel = $statusLabel

# Создание кнопки экспорта в CSV
$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Location = New-Object System.Drawing.Point(610, 560)
$buttonExport.Size = New-Object System.Drawing.Size(100, 40)
$buttonExport.Text = "Экспорт в CSV"
$buttonExport.Enabled = $false
$form.Controls.Add($buttonExport)
$syncHash.ButtonExport = $buttonExport

# Добавление кнопки "О программе"
$buttonAbout = New-Object System.Windows.Forms.Button
$buttonAbout.Location = New-Object System.Drawing.Point(720, 560)
$buttonAbout.Size = New-Object System.Drawing.Size(50, 40)
$buttonAbout.Text = "?"
$buttonAbout.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonAbout)
$syncHash.ButtonAbout = $buttonAbout

# Настройка таймера для проверки состояния выполнения задач
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$syncHash.Timer = $timer

# Создание RunspacePool для выполнения параллельных задач
$syncHash.RunspacePool = [runspacefactory]::CreateRunspacePool(1, 5)
$syncHash.RunspacePool.ApartmentState = "STA"
$syncHash.RunspacePool.Open()

# Функция для записи в лог из основного потока
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $message"
    $syncHash.LogTextBox.AppendText("$logMessage`r`n")
    $syncHash.LogTextBox.ScrollToCaret()
    Write-Host $logMessage
}

# Функция для записи в лог из другого потока
function Write-LogFromRunspace {
    param ([string]$message)
    $syncHash.Form.Invoke([Action]{
        $timestamp = Get-Date -Format "HH:mm:ss"
        $logMessage = "[$timestamp] $message"
        $syncHash.LogTextBox.AppendText("$logMessage`r`n")
        $syncHash.LogTextBox.ScrollToCaret()
        Write-Host $logMessage
    })
}

# Функция для отображения информации о приложении
$buttonAbout.Add_Click({
    # Создание формы "О программе"
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "О программе"
    $aboutForm.Size = New-Object System.Drawing.Size(450, 250)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false
    
    # Название программы
    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Location = New-Object System.Drawing.Point(20, 20)
    $labelTitle.Size = New-Object System.Drawing.Size(400, 30)
    $labelTitle.Text = "AD Group Finder"
    $labelTitle.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $aboutForm.Controls.Add($labelTitle)
    
    # Автор
    $labelAuthor = New-Object System.Windows.Forms.Label
    $labelAuthor.Location = New-Object System.Drawing.Point(20, 60)
    $labelAuthor.Size = New-Object System.Drawing.Size(400, 25)
    $labelAuthor.Text = "Автор: Virsuløn DΞv"
    $labelAuthor.Font = New-Object System.Drawing.Font("Arial", 10)
    $aboutForm.Controls.Add($labelAuthor)
    
    # Версия
    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Location = New-Object System.Drawing.Point(20, 90)
    $labelVersion.Size = New-Object System.Drawing.Size(400, 25)
    $labelVersion.Text = "Версия: 0.0.05.1 alpha"
    $labelVersion.Font = New-Object System.Drawing.Font("Arial", 10)
    $aboutForm.Controls.Add($labelVersion)
    
    # Сайт (кликабельная ссылка)
    $linkLabel = New-Object System.Windows.Forms.LinkLabel
    $linkLabel.Location = New-Object System.Drawing.Point(20, 120)
    $linkLabel.Size = New-Object System.Drawing.Size(400, 25)
    $linkLabel.Text = "github.com/G3XQSOFT-AG/ADGroupFinder"
    $linkLabel.LinkColor = [System.Drawing.Color]::Blue
    $linkLabel.ActiveLinkColor = [System.Drawing.Color]::Red
    $linkLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $aboutForm.Controls.Add($linkLabel)
    
    # Событие для открытия ссылки по клику
    $linkLabel.Add_LinkClicked({
        [System.Diagnostics.Process]::Start("https://github.com/G3XQSOFT-AG/ADGroupFinder")
    })
    
    # Кнопка "OK" для закрытия окна
    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Location = New-Object System.Drawing.Point(180, 170)
    $buttonOK.Size = New-Object System.Drawing.Size(80, 30)
    $buttonOK.Text = "OK"
    $buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $aboutForm.Controls.Add($buttonOK)
    $aboutForm.AcceptButton = $buttonOK
    
    # Отображение окна с информацией об авторе
    [void]$aboutForm.ShowDialog()
})

# Функция для проверки состояния задач
$timer.Add_Tick({
    if ($syncHash.Jobs.Count -gt 0) {
        $completedJobs = $syncHash.Jobs | Where-Object { $_.AsyncResult.IsCompleted }
        
        foreach ($job in $completedJobs) {
            try {
                # Получаем результат выполнения задачи
                $result = $job.PowerShell.EndInvoke($job.AsyncResult)
                $syncHash.SearchResults = $result
                $syncHash.SearchCompleted = $true
                $syncHash.Error = $null
            } catch {
                $syncHash.Error = $_.Exception.Message
                $syncHash.SearchCompleted = $true
            } finally {
                # Очистка ресурсов
                $job.PowerShell.Dispose()
                $syncHash.Jobs.Remove($job)
            }
        }
        
        # Если все задачи завершены
        if ($syncHash.Jobs.Count -eq 0) {
            $timer.Stop()
            
            # Скрытие индикатора прогресса
            $syncHash.ProgressBar.MarqueeAnimationSpeed = 0
            $syncHash.ProgressBar.Visible = $false
            $syncHash.DataGridView.Visible = $true
            
            # Если произошла ошибка
            if ($syncHash.Error) {
                $syncHash.StatusLabel.Text = "Ошибка при поиске: $($syncHash.Error)"
                Write-Log "Ошибка при поиске: $($syncHash.Error)"
                [System.Windows.Forms.MessageBox]::Show("Произошла ошибка при поиске: $($syncHash.Error)", 
                    "Ошибка поиска", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $syncHash.ButtonSearch.Enabled = $true
                return
            }
            
            # Обработка результатов поиска
            $results = $syncHash.SearchResults
            
            if ($results -and ($results.Count -gt 0)) {
                # Конвертация результатов в формат для отображения
                $dataTable = New-Object System.Data.DataTable
                $dataTable.Columns.Add("Имя группы")
                $dataTable.Columns.Add("SamAccountName")
                $dataTable.Columns.Add("Домен")
                $dataTable.Columns.Add("Описание")
                $dataTable.Columns.Add("Информация")
                
                foreach ($result in $results) {
                    $row = $dataTable.NewRow()
                    $row["Имя группы"] = $result.Name
                    $row["SamAccountName"] = $result.samaccountname
                    $row["Домен"] = $result.Domain
                    $row["Описание"] = $result.Description
                    $row["Информация"] = $result.info
                    $dataTable.Rows.Add($row)
                }
                
                $syncHash.DataGridView.DataSource = $dataTable
                $syncHash.ButtonExport.Enabled = $true
                $syncHash.StatusLabel.Text = "Найдено групп: " + $results.Count
                Write-Log "Поиск завершен. Найдено групп: $($results.Count)"
            } else {
                $syncHash.StatusLabel.Text = "Группы не найдены"
                Write-Log "Поиск завершен. Группы не найдены."
                [System.Windows.Forms.MessageBox]::Show("По указанному шаблону группы не найдены. Попробуйте изменить параметры поиска.", 
                    "Группы не найдены", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            
            # Возвращаем кнопку поиска в активное состояние
            $syncHash.ButtonSearch.Enabled = $true
        }
    }
})

# Функция для отображения информации о домене
$buttonDomainInfo.Add_Click({
    try {
        $selectedDomain = $comboDomain.SelectedItem
        if ($selectedDomain -eq "Все домены" -or $selectedDomain -eq "Ошибка получения доменов") {
            [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите конкретный домен для получения информации", 
                "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        $domainInfo = Get-ADDomain -Identity $selectedDomain
        $infoText = @"
Информация о домене: $($domainInfo.DNSRoot)

NETBIOS имя: $($domainInfo.NetBIOSName)
Режим домена: $($domainInfo.DomainMode)
Родительский домен: $($domainInfo.ParentDomain)
Дочерние домены: $($domainInfo.ChildDomains -join ", ")
Контроллеры: $($domainInfo.ReplicaDirectoryServers -join ", ")

Forest Root Domain: $($domainInfo.Forest)
RID Master: $($domainInfo.RIDMaster)
PDC Emulator: $($domainInfo.PDCEmulator)
"@
        
        [System.Windows.Forms.MessageBox]::Show($infoText, "Информация о домене $($domainInfo.DNSRoot)", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при получении информации о домене: $($_.Exception.Message)", 
            "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Функция для проверки подключения
$buttonTestConnection.Add_Click({
    try {
        $selectedDomain = $comboDomain.SelectedItem
        if ($selectedDomain -eq "Все домены" -or $selectedDomain -eq "Ошибка получения доменов") {
            [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите конкретный домен для проверки подключения", 
                "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        
        $domainInfo = Get-ADDomain -Identity $selectedDomain
        $dc = $domainInfo.PDCEmulator
        
        Write-Log "Проверка подключения к DC: $dc"
        
        if (Test-Connection -ComputerName $dc -Count 1 -Quiet) {
            # Проверяем доступность AD
            $testUser = Get-ADUser -Filter "Name -like '*'" -ResultSetSize 1 -Server $dc -ErrorAction Stop
            $testGroup = Get-ADGroup -Filter "Name -like '*'" -ResultSetSize 1 -Server $dc -ErrorAction Stop
            
            [System.Windows.Forms.MessageBox]::Show("Успешное подключение к домену $selectedDomain`nКонтроллер домена: $dc`nНайдена группа: $($testGroup.Name)`nНайден пользователь: $($testUser.Name)", 
                "Подключение проверено", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Не удалось подключиться к контроллеру домена: $dc", 
                "Ошибка подключения", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при проверке подключения: $($_.Exception.Message)", 
            "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Функция для выполнения поиска
$buttonSearch.Add_Click({
    $searchPattern = $textBoxSearch.Text
    $selectedDomain = $comboDomain.SelectedItem
    
    # Очистка предыдущих результатов
    $dataGridView.DataSource = $null
    $buttonExport.Enabled = $false
    
    # Проверка наличия шаблона поиска
    if ([string]::IsNullOrWhiteSpace($searchPattern)) {
        $statusLabel.Text = "Ошибка: Введите шаблон для поиска"
        Write-Log "Ошибка: Шаблон поиска не указан"
        return
    }
    
    # Отображение индикатора прогресса
    $progressBar.Visible = $true
    $progressBar.MarqueeAnimationSpeed = 30
    $dataGridView.Visible = $false
    
    # Блокировка кнопки поиска во время выполнения
    $buttonSearch.Enabled = $false
    
    $statusLabel.Text = "Выполняется поиск групп с шаблоном: $searchPattern"
    
    $methodDescription = if ($radioStandard.Checked) {
        'Стандартный'
    } elseif ($radioLDAP.Checked) {
        'LDAP'
    } else {
        'Расширенный'
    }
    
    Write-Log "Начат поиск групп с шаблоном: $searchPattern, Домен: $selectedDomain, Метод: $methodDescription"
    
    # Создание PowerShell скрипта для выполнения в отдельном потоке
    $scriptBlock = {
        param($searchPattern, $selectedDomain, $domains, $isStandard, $isLDAP, $isAdvanced)
        
        # Функция для записи в лог
        function Write-LogFromThread {
            param ([string]$message)
            # Запись в лог (в потоке будет игнорироваться, но полезно для отладки)
            Write-Host $message
        }
        
        # Функция для поиска групп AD
        function Search-ADGroups {
            param (
                [string]$Domain,
                [string]$SearchPattern,
                [bool]$IsStandard,
                [bool]$IsLDAP,
                [bool]$IsAdvanced
            )
            
            $groups = @()
            
            # Определяем, нужно ли явно указывать сервер
            $serverParam = @{}
            if ($Domain -ne "Ошибка получения доменов") {
                $serverParam = @{ Server = $Domain }
            }
            
            # Формируем base DN для поиска
            $searchBase = ""
            try {
                $domainObj = Get-ADDomain $Domain
                $searchBase = $domainObj.DistinguishedName
            } catch {
                Write-LogFromThread "Не удалось получить DN домена $Domain. Ошибка: $($_.Exception.Message)"
                # Продолжаем без searchBase
            }
            
            $searchBaseParam = @{}
            if ($searchBase) {
                $searchBaseParam = @{ SearchBase = $searchBase }
            }
            
            try {
                # Выбор метода поиска
                if ($IsStandard) {
                    # Стандартный метод - используем строковый фильтр
                    $filter = "Name -like '$SearchPattern'"
                    Write-LogFromThread "Выполнение стандартного поиска с фильтром: $filter в домене $Domain"
                    
                    $adGroups = Get-ADGroup -Filter $filter -Properties Description,info @serverParam @searchBaseParam
                    
                    if ($adGroups) {
                        foreach ($group in $adGroups) {
                            $groupInfo = [PSCustomObject]@{
                                Name = $group.Name
                                samaccountname = $group.samaccountname
                                Domain = $Domain
                                Description = $group.Description
                                info = $group.info
                            }
                            $groups += $groupInfo
                        }
                    }
                }
                elseif ($IsLDAP) {
                    # LDAP-фильтр
                    $ldapFilter = "(name=$SearchPattern)"
                    Write-LogFromThread "Выполнение LDAP поиска с фильтром: $ldapFilter в домене $Domain"
                    
                    $adGroups = Get-ADGroup -LDAPFilter $ldapFilter -Properties Description,info @serverParam @searchBaseParam
                    
                    if ($adGroups) {
                        foreach ($group in $adGroups) {
                            $groupInfo = [PSCustomObject]@{
                                Name = $group.Name
                                samaccountname = $group.samaccountname
                                Domain = $Domain
                                Description = $group.Description
                                info = $group.info
                            }
                            $groups += $groupInfo
                        }
                    }
                }
                else {
                    # Расширенный метод - используем Get-ADObject с более гибкими параметрами
                    Write-LogFromThread "Выполнение расширенного поиска для $SearchPattern в домене $Domain"
                    
                    # Поиск через Get-ADObject
                    $adObjects = Get-ADObject -Filter "objectClass -eq 'group' -and name -like '$SearchPattern'" -Properties Description,info,samaccountname @serverParam @searchBaseParam
                    
                    if ($adObjects) {
                        foreach ($obj in $adObjects) {
                            $groupInfo = [PSCustomObject]@{
                                Name = $obj.Name
                                samaccountname = $obj.samAccountName
                                Domain = $Domain
                                Description = $obj.Description
                                info = $obj.info
                            }
                            $groups += $groupInfo
                        }
                    }
                    
                    # Дополнительный метод поиска через Where-Object
                    if ($groups.Count -eq 0) {
                        $allGroups = Get-ADGroup -Filter * -Properties Description,info @serverParam @searchBaseParam
                        $filteredGroups = $allGroups | Where-Object { $_.Name -like $SearchPattern }
                        
                        if ($filteredGroups) {
                            foreach ($group in $filteredGroups) {
                                $groupInfo = [PSCustomObject]@{
                                    Name = $group.Name
                                    samaccountname = $group.samaccountname
                                    Domain = $Domain
                                    Description = $group.Description
                                    info = $group.info
                                }
                                $groups += $groupInfo
                            }
                        }
                    }
                }
            }
            catch {
                Write-LogFromThread "Ошибка при поиске в домене ${Domain}: $($_.Exception.Message)"
                # Продолжаем работу с другими доменами, но сохраняем информацию об ошибке
            }
            
            return $groups
        }
        
        # Основной код выполнения поиска
        $allResults = @()
        
        try {
            if ($selectedDomain -eq "Все домены") {
                foreach ($domain in $domains) {
                    if ($domain -eq "Все домены" -or $domain -eq "Ошибка получения доменов") { continue }
                    
                    $domainResults = Search-ADGroups -Domain $domain -SearchPattern $searchPattern -IsStandard $isStandard -IsLDAP $isLDAP -IsAdvanced $isAdvanced
                    if ($domainResults) {
                        $allResults += $domainResults
                    }
                }
            } else {
                $domainResults = Search-ADGroups -Domain $selectedDomain -SearchPattern $searchPattern -IsStandard $isStandard -IsLDAP $isLDAP -IsAdvanced $isAdvanced
                if ($domainResults) {
                    $allResults += $domainResults
                }
            }
            
            return $allResults
            
        } catch {
            throw "Ошибка при выполнении поиска: $($_.Exception.Message)"
        }
    }
    
    # Создание PowerShell объекта для выполнения в RunspacePool
    $powershell = [powershell]::Create()
    $powershell.RunspacePool = $syncHash.RunspacePool
    
    # Добавление скрипта и параметров
    [void]$powershell.AddScript($scriptBlock)
    [void]$powershell.AddArgument($searchPattern)
    [void]$powershell.AddArgument($selectedDomain)
    [void]$powershell.AddArgument($syncHash.Domains)
    [void]$powershell.AddArgument($radioStandard.Checked)
    [void]$powershell.AddArgument($radioLDAP.Checked)
    [void]$powershell.AddArgument($radioAdvanced.Checked)
    
    # Начало асинхронного выполнения
    $asyncResult = $powershell.BeginInvoke()
    
    # Сохранение задачи для последующей проверки
    [void]$syncHash.Jobs.Add([PSCustomObject]@{
        PowerShell = $powershell
        AsyncResult = $asyncResult
    })
    
    # Запуск таймера для проверки выполнения задачи
    $timer.Start()
})

# Функция для экспорта результатов в CSV
$buttonExport.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV файлы (*.csv)|*.csv"
    $saveFileDialog.Title = "Сохранить результаты как CSV"
    $saveFileDialog.FileName = "AD_Groups_Export.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        try {
            # Получаем данные из DataGridView и сохраняем в CSV
            $dt = $dataGridView.DataSource
            $data = @()
            foreach ($row in $dt.Rows) {
                $obj = New-Object PSObject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Имя группы" -Value $row."Имя группы"
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "SamAccountName" -Value $row."SamAccountName"
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Домен" -Value $row."Домен"
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Описание" -Value $row."Описание"
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Информация" -Value $row."Информация"
                $data += $obj
            }
            
            $data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            $statusLabel.Text = "Данные успешно экспортированы в $filePath"
            Write-Log "Данные успешно экспортированы в $filePath"
        } catch {
            $statusLabel.Text = "Ошибка при экспорте: " + $_.Exception.Message
            Write-Log "Ошибка при экспорте: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Произошла ошибка при экспорте: $($_.Exception.Message)", 
                "Ошибка экспорта", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Очистка ресурсов при закрытии формы
$form.Add_FormClosing({
    # Остановка таймера
    $timer.Stop()
    
    # Остановка и удаление всех запущенных заданий
    foreach ($job in $syncHash.Jobs) {
        try {
            $job.PowerShell.Stop()
            $job.PowerShell.Dispose()
        } catch {
            # Игнорируем ошибки при очистке
        }
    }
    
    # Закрытие RunspacePool
    try {
        $syncHash.RunspacePool.Close()
        $syncHash.RunspacePool.Dispose()
    } catch {
        # Игнорируем ошибки при очистке
    }
})

# Инициализация журнала
Write-Log "Приложение запущено. Модуль Active Directory загружен."
Write-Log "Текущий домен: $currentDomain"

# Отображение формы
[void]$form.ShowDialog()
