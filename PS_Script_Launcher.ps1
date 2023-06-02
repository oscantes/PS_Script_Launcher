Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing, PresentationFramework

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PSScript Launcher'
$form.Size = New-Object System.Drawing.Size(745,577) #lonng way of defining size etc.
$form.AutoSize = $true
$form.StartPosition = 'CenterScreen'
$form.BackColor = "#b2d5e5"
$form.AllowDrop = $true
#$form.TopMost = $true

$tabcontrol = New-object System.Windows.Forms.TabControl
$tabcontrol.Location = "20,20"
$tabcontrol.Size = "690,500"
#$tabcontrol.BackColor = "#b2d5e5"
$tabcontrol.Add_DoubleClick({ Enable-Experimental_Tab})
$form.Controls.Add($tabcontrol)

#$tab_scriptlauncher.DataBindings.DefaultDataSourceUpdateMode = 0 
$tab_scriptlauncher = New-object System.Windows.Forms.Tabpage
$tab_scriptlauncher.Text = "BTM"
$tab_scriptlauncher.BackColor = "lightyellow"
$tabcontrol.Controls.Add($tab_scriptlauncher)

function Enable-Experimental_Tab {
    $tab_experimental = New-object System.Windows.Forms.Tabpage
    $tab_experimental.Text = "*Experimental*"
    $tab_experimental.BackColor = "LightGreen"
    $tabcontrol.Controls.Add($tab_experimental)

    $start_script_button = New-Object System.Windows.Forms.Button
    $start_script_button.Location = New-Object System.Drawing.Point(360,200)
    $start_script_button.Size = New-Object System.Drawing.Size(240,60)
    $start_script_button.Text = 'Start Script'
    $start_script_button.BackColor = '#EEEE55'
    $start_script_button.Add_Click($button_click)
    $tab_experimental.Controls.Add($start_script_button)
}

$scriptpicker_label = New-Object System.Windows.Forms.Label
$scriptpicker_label.Location = New-Object System.Drawing.Point(20,20)
$scriptpicker_label.AutoSize = $true
$scriptpicker_label.Text = 'Select a script from the list > '
$tab_scriptlauncher.Controls.Add($scriptpicker_label)

#multiple selection listbox'a döndürülebilir
$script_picker = New-Object System.Windows.Forms.ComboBox
$script_picker.Location  = New-Object System.Drawing.Point(20,40)
$script_picker.Width = 380
[void]$script_picker.Items.Add('create_mistlog')
[void]$script_picker.Items.Add('delete_ANKA_temp')
[void]$script_picker.Items.Add('delete_jabber_temp')
[void]$script_picker.Items.Add('get_OS_info')
[void]$script_picker.Items.Add('application_Launcher')
[void]$script_picker.Items.Add('get_Java_version')
[void]$script_picker.Items.Add('active_directory_VPN_query')
[void]$script_picker.Items.Add('TEST_runspace')
[void]$script_picker.Items.Add('TEST_get_computername_from_DW')
[void]$script_picker.Items.Add('TEST_Query_session')
$script_picker.Add_SelectedIndexChanged{Start-Visualizer}
$tab_scriptlauncher.Controls.Add($script_picker)

$computername_label = New-Object System.Windows.Forms.Label
$computername_label.Location = New-Object System.Drawing.Point(440,20)
$computername_label.AutoSize = $true
$computername_label.Text = 'Enter computername/IP or double click(DW)'
$tab_scriptlauncher.Controls.Add($computername_label)

$computernamebox = New-Object System.Windows.Forms.TextBox
$computernamebox.Location = New-Object System.Drawing.Point(440,40)
$computernamebox.Size = New-Object System.Drawing.Size(220,100)
$computernamebox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { Start-SelectedScript -script_name $script_picker.SelectedItem}})
$computernamebox.Add_DoubleClick({ Start-TEST_get_computername_from_DW })
$tab_scriptlauncher.Controls.Add($computernamebox)

$username_label = New-Object System.Windows.Forms.Label
$username_label.Location = New-Object System.Drawing.Point(440,80)
$username_label.AutoSize = $true
$username_label.Text = 'Enter username'
$tab_scriptlauncher.Controls.Add($username_label)

$usernamebox = New-Object System.Windows.Forms.TextBox
$usernamebox.Location = New-Object System.Drawing.Point(440,100)
$usernamebox.Size = New-Object System.Drawing.Size(220,100)
$usernamebox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { Start-SelectedScript -script_name $script_picker.SelectedItem}})
$usernamebox.Add_DoubleClick({ Start-TEST_Query_session })
$tab_scriptlauncher.Controls.Add($usernamebox)

$button_click = {Start-SelectedScript -script_name $script_picker.SelectedItem}
#$button_click = {Start-loading_animation}
$start_script_button = New-Object System.Windows.Forms.Button
$start_script_button.Location = New-Object System.Drawing.Point(440,180)
$start_script_button.Size = New-Object System.Drawing.Size(220,60)
$start_script_button.Text = 'Start Script'
$start_script_button.BackColor = '#EEEE55'
$start_script_button.Enabled = $false
$start_script_button.Add_Click($button_click)
$tab_scriptlauncher.Controls.Add($start_script_button)

$infobox_label = New-Object System.Windows.Forms.Label
$infobox_label.Location = New-Object System.Drawing.Point(20,80)
$infobox_label.AutoSize = $true
$infobox_label.Text = 'Infobox'
$tab_scriptlauncher.Controls.Add($infobox_label)

$infobox = New-Object System.Windows.Forms.TextBox
$infobox.Multiline = $true
$infobox.Location = New-Object System.Drawing.Point(20,100)
$infobox.Size = New-Object System.Drawing.Size(380,140)
$infobox.AllowDrop = $true
$infobox.ReadOnly = $true
$infobox.Add_DoubleClick({ Start-get_OS_info })
$tab_scriptlauncher.controls.Add($infobox)

$logbox_label = New-Object System.Windows.Forms.Label
$logbox_label.Location = New-Object System.Drawing.Point(20,260)
$logbox_label.AutoSize = $true
$logbox_label.Text = 'Logbox'
$tab_scriptlauncher.Controls.Add($logbox_label)

#sort/show items in descending order
$logbox = New-Object System.Windows.Forms.ListBox
$logbox.Location = New-Object System.Drawing.Point(20,280)
$logbox.Size = New-Object System.Drawing.Size(640,180)

$tab_scriptlauncher.controls.Add($logbox)

function Start-Visualizer {
    #switch kullanılsa? ama listbox olursa(multiple selection) switch çalışmaz
    #$param ile çalışacak hale getirilebilir, default ayarlar olur ve sadece modifiye edilecekler param olarak iletilir
        $start_script_button.Enabled = $true

    if ($script_picker.SelectedItem -eq 'create_mistlog') {
        $computernamebox.BackColor = "#EEEE55"
        $usernamebox.BackColor  = "window"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Zips files of C:\MISTLOG folder on remote computer"
    }
    elseif ($script_picker.SelectedItem -eq 'delete_ANKA_temp') {
        $computernamebox.BackColor  = "#EEEE55"
        $usernamebox.BackColor  = "#EEEE55"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Deletes \AppData\Local\Apps\2.0 and Z:\Belgelerim\username\My Documents\Nar"
    }
    elseif ($script_picker.SelectedItem -eq 'delete_jabber_temp') {
        $computernamebox.BackColor  = "#EEEE55"
        $usernamebox.BackColor  = "#EEEE55"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Kills all Jabber processes and deletes cache files"
    }
    elseif ($script_picker.SelectedItem -eq 'get_OS_info') {
        $computernamebox.BackColor  = "#EEEE55"
        $usernamebox.BackColor  = "window"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Gets OS info of computer(double clicking on infobox also works)"
    }
    elseif ($script_picker.SelectedItem -eq 'application_Launcher') {
        $computernamebox.BackColor  = "#EEEE55"
        $usernamebox.BackColor  = "window"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Kills all Internet Explorer processes and deletes history?? & cache files`r
(A user must be logged in, doesn't work with RDP connection)"
    }
    elseif ($script_picker.SelectedItem -eq 'get_Java_version') {
        $computernamebox.BackColor  = "#EEEE55"
        $usernamebox.BackColor  = "window"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Shows versions of installed Java programs`r
(Takes more time than usual)"
    }
    elseif ($script_picker.SelectedItem -eq 'active_directory_VPN_query') {
        $computernamebox.BackColor  = "window"
        $usernamebox.BackColor  = "#EEEE55"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Queries if user is in SSLVPN group or not"
    }

    elseif ($script_picker.SelectedItem -eq 'TEST_runspace') {
        $computernamebox.BackColor  = "window"
        $usernamebox.BackColor  = "window"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Tests runspace"
    }
    elseif ($script_picker.SelectedItem -eq 'TEST_get_computername_from_DW') {
        $computernamebox.BackColor  = "window"
        $usernamebox.BackColor  = "window"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "Tests if it can get the computername from Dameware MRC"
    }
    elseif ($script_picker.SelectedItem -eq 'TEST_Query_session') {
        $computernamebox.BackColor  = "#EEEE55"
        $usernamebox.BackColor  = "#EEEE55"
        $infobox.BackColor = "window"
        $infobox.TextAlign = "left"
        $infobox.Text = "if username is not entered.."
    }
}

function Start-SelectedScript {
    param ($script_name)

    $logbox_date = (get-date -format "dd.MM.yyyy HH:mm:ss")
    $logbox.Items.Add($logbox_date + " - " + $computernamebox.Text+" için "+$script_name+" çalıştırılıyor")

    Start-loading_animation("start")
    
    #checks if first letter of computernamebox is hostname or ip; if it's ip, converts it to hostname
    #if computername is not enabled, meaning there is no need to computername, it passes this control
    if ($computernamebox.Text -match '^[a-z]' -or $computernamebox.BackColor -eq "window") {
        Invoke-Command -ScriptBlock (get-item Function:\'Start-'$script_name).ScriptBlock
    
    }
    elseif ($computernamebox.Text -match '^[0-9]') { #second check if it could get it from DW MRC
        $infobox.Text = "Hostname isn't entered, trying to find hostname from IP"
        #alınacak hatayı önlemek için ping'lenebiliyor mu kontrolü gir buraya
        $hostname = [System.Net.Dns]::GetHostByAddress($computernamebox.Text).Hostname
        $computernamebox.Text = $hostname
        Invoke-Command -ScriptBlock (get-item Function:\'Start-'$script_name).ScriptBlock
        
    }
    else {
        $infobox.Text = "computername or username is needed but it's null, process will exit"
        Start-loading_animation("error")
        return
    }

    $logbox_date = (get-date -format "dd.MM.yyyy HH:mm:ss")
    $logbox.Items.Add($logbox_date + " - " + $computernamebox.Text+" için "+$script_name+" çalıştı")

    Start-loading_animation("stop")

}

function Start-loading_animation {
    param ($x)
    switch ($x) {
        start {
            $start_script_button.Text = ''
            $load_image = [System.Drawing.Image]::FromFile("$PSScriptRoot\PS_Script_Launcher_media_please_wait2.png")
            $start_script_button.Image = $load_image
            $form.Update()
          }
        stop {
            $load_image = [System.Drawing.Image]::FromFile("$PSScriptRoot\PS_Script_Launcher_media_job_done.png")
            $start_script_button.Image = $load_image
            $form.Update()
            Start-Sleep -Seconds 1
            $start_script_button.Image = $null
            $start_script_button.Text = 'Start Script'
            $form.Update()
        }
        error {
            $load_image = [System.Drawing.Image]::FromFile("$PSScriptRoot\PS_Script_Launcher_media_an_error_occured.png")
            $start_script_button.Image = $load_image
            $form.Update()
            Start-Sleep -Seconds 1
            $start_script_button.Image = $null
            $start_script_button.Text = 'Start Script'
            $form.Update()
        }
        Default {
            Write-Host "Animasyonda hata"
        }
    }    
}



#buradan aşağısı listeden seçilen script fonksiyonları

function Start-create_mistlog {

    Invoke-Command -ComputerName $computernamebox.Text -ScriptBlock{  #credentials?
        if (Test-Path -path Z:\MISTLOG) {Remove-Item Z:\MISTLOG -recurse}
        New-Item Z:\MISTLOG -Type Directory
        get-childitem C:\MISTLOG | where-Object LastWriteTime -GT (Get-Date).Date | copy-item -destination Z:\MISTLOG
        $logfilename = "Z:\MISTLOG\MISTLOG_"+(hostname)+"_"+(get-date -format "dd_MM_yyyy_HH_mm_ss") #don't use "." or ":"
        Compress-Archive Z:\MISTLOG\* -Update -DestinationPath $logfilename
        Remove-Item -path Z:\MISTLOG -Exclude *.zip -Recurse
    }
    # dosyayı kendi bilgisayarımıza alıyoruz
    $session = New-PSSession -ComputerName $computernamebox.Text
    Copy-Item Z:\MISTLOG\* -Destination "Z:\belgelerim\b7lwbto\My Documents\Desktop" -FromSession $session -recurse
    Invoke-Command -Session $session -Command {Remove-Item Z:\MISTLOG -recurse}
    $infobox.BackColor = "lightgreen"

}

function Start-delete_ANKA_temp {

    Invoke-Command -ComputerName $computernamebox.Text -ScriptBlock {
        taskkill /F /IM Anka.Shell.exe;
        taskkill /F /IM Nar.exe;
        taskkill /F /IM EZE2IMG.EXE;

        #aşağısı invoke içinde çalışacak mı kontrol et olmazsa invoke dışına alıp dizini güncelle
        $path1 = "C:\Users\"+$using:usernamebox.Text+"\AppData\Local\Apps\2.0"
        $path2 = "Z:\Belgelerim\"+$using:usernamebox.Text+"\My Documents\Nar"

        Remove-Item $path1, $Path2 -Recurse -Force
        #if (Test-Path -path $path1){write-host "\AppData\Local\Apps\2.0 klasörü bulundu"; Remove-Item $path1 -Recurse -Force}
    }

}

function Start-delete_jabber_temp {
    $invoke_results = Invoke-Command -ComputerName $computernamebox.Text -ScriptBlock{
        
        $processes = Get-WmiObject -Class Win32_Process -Filter "name='ciscojabber.exe'"
        $count1 = 0
        foreach($process in $processes){   #terminate edemeyince hata vermeyecek şekilde ayarla
            if($process.terminate().returnvalue -eq 0){$count1++}}

        $username = $using:usernamebox.Text
        $path1 = "C:\Users\$username\AppData\Local\Cisco\Unified Communications\Jabber"
        $path2 = "C:\Users\$username\AppData\Roaming\Cisco\Unified Communications\Jabber"
        $path3 = "C:\ProgramData\Cisco Systems\Cisco Jabber"

        $count2 = 0
        if (Test-Path -path $path1){Remove-Item $path1 -Recurse; $count2++}
        if (Test-Path -path $path2){Remove-Item $path2 -Recurse; $count2++}
        if (Test-Path -path $path3){Remove-Item $path3 -Recurse -Force; $count2++}
            
        $count1 #returns terminated processes
        $count2 #returns deleted paths
    }
    $logbox_date = (get-date -format "dd.MM.yyyy HH:mm:ss")
    $logbox.Items.Add($logbox_date + " - " + $invoke_results[0].toString() + " adet işlem sonlandırıldı, " + $invoke_results[1].toString() + " adet dizin silindi")
}

function Start-get_OS_info { #bu fonksiyon yavaş çalışmasın diye ping testi eklendi

    if (Test-Connection $computernamebox.Text -Quiet) {

        $OS_Computername = $computernamebox.Text
        $OS_Caption = (Get-WmiObject Win32_OperatingSystem -Computername $computernamebox.Text).Caption
        $OS_BuildNumber = (Get-WmiObject Win32_OperatingSystem -Computername $computernamebox.Text).BuildNumber
        $OS_InstallDate = (Get-WmiObject Win32_OperatingSystem).ConvertToDateTime( (Get-WmiObject Win32_OperatingSystem -Computername $computernamebox.Text).InstallDate).toString('dd.MM.yyyy, HH:mm:ss')
        $OS_LastBootTime = (Get-WmiObject Win32_OperatingSystem).ConvertToDateTime( (Get-WmiObject Win32_OperatingSystem -Computername $computernamebox.Text).LastBootUpTime).toString('dd.MM.yyyy, HH:mm:ss')
    
        $infobox.Text = "OS Info:`r
            Computername: $OS_Computername`r
            Caption: $OS_Caption`r
            Build Number: $OS_BuildNumber`r
            Install Date: $OS_InstallDate`r
            Last Boot Time: $OS_LastBootTime"
    }
    else {
        $infobox.Text = "Sistem bilgileri getirilemedi, bilgisayar kapalı olabilir"
    }

}

#aşağıdaki fonksiyonun olayı ne? gerekirse sil yada notes kısmına al, junk gibi duruyor
function Start-application_Launcher {
<#
query session zaman kaybı + psexec'te geçen süre, direkt start process ile yapılamaz mı sessionid onda yok mu
eğer böyle olursa hepsi invoke command'e atılabilir
#>
    if ((Test-Path -path "C:\Windows\System32\PsExec.exe") -eq $false){
        Copy-Item \\NB341127C2020\Z$\PsExec.exe -Destination "C:\Windows\System32\PsExec.exe"}
    if ((Test-Path -path "C:\Windows\System32\PsExec.exe") -eq $false){
        $infobox.Text = "Bu script'in çalışması için gerekli yardımcı program bilgisayarınızda bulunmuyor,`r
Kopyalama denendi fakat işlem başarılı olmadı,`r
(0x23482785) hata kodunu geliştirici ile paylaşabilirsiniz."
    }
    else {

        #yeni fonksiyonu gir yeni fonksiyon derken?

        $param = "\\"+$computernamebox.Text
        psexec $param /accepteula -s -i $sessionid cmd.exe /C "notepad";
    }
}

function Start-get_Java_version {
    #java installation date?
    #sorgulanacak programları listbox'ta seçtirmek?
    $programs = Get-WmiObject Win32_Product -ComputerName $computernamebox.Text -Filter "Name LIKE '%Java%'" | Select-Object Name, Version
    foreach($program in $programs){$logbox.Items.Add($program)}

    # $asd = Invoke-Command -ComputerName $computernamebox.Text -ScriptBlock{
    #     $asd2 = Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | where-object DisplayName -eq 'Cisco Jabber'
    #     return $asd2.InstallDate
    # }
    # $logbox.Items.Add("Cisco Jabber install date > $asd")
}



function Start-active_directory_VPN_query {
    if ($usernamebox.Text -eq "") {
        $infobox.Text = "username girilmedi, process will exit"
    }
    else {
        $ADuserGroups = Get-ADPrincipalGroupMembership $usernamebox.Text
        #filter group before out-string, takes too much time this way
        $strAduserGroups = $ADuserGroups | Out-String
        if(($strAduserGroups).Contains("PYSSLVPNAccountBaseGRP")){
            $infobox.Text = $usernamebox.Text+": VPN yetkisi VAR"}
        else{$infobox.Text = $usernamebox.Text+": VPN yetkisi YOK"}
}

}

function Start-TEST_runspace {

    $Runspace = [runspacefactory]::CreateRunspace()
    $PowerShell1 = [powershell]::Create()
    $PowerShell1.runspace = $Runspace
    $Runspace.Open()
    
    [void]$PowerShell1.AddScript({
        #these don't seem to reach out until EndInvoke()
        Get-Date
        Start-Sleep -Seconds 5
    })

    $AsyncObject = $PowerShell1.BeginInvoke() #"begin" to better monitor the state
    $PowerShell1.EndInvoke($AsyncObject)
    $Data = $PowerShell1.EndInvoke($AsyncObject) #to use the output
    write-host $Data
    $PowerShell1.Dispose()
}

function Start-TEST_get_computername_from_DW {

    try {
    #birden çok dw penceresi olursa?
    $getasd = Get-Process DWRCC -ErrorAction 'Stop' #erroraction seems to be the key for this to work properly
    $gotasd = $getasd.MainWindowTitle
    $separator = " "
    $dw_comp_name = $gotasd.split($separator)
    $logbox_date = (get-date -format "dd.MM.yyyy HH:mm:ss")
    $logbox.Items.Add($logbox_date + " - " + $dw_comp_name[0] + " bilgisayar adı elde edildi")
    $computernamebox.Text = $dw_comp_name[0]
    }
    catch {
        $infobox.Text = "Bilgisayar adı elde edilemedi, Dameware MRC kapalı olabilir"
        $infobox.BackColor = "#FFAAAA"
    }
        #buraya neden giriyor? infobox rengi değişmesine rağmen text'i yazdırmıyor, neden?
        #Job is not done properly but gui says job done, does this need to be fixed?
    }

    function Start-TEST_Query_session {
        $param2 = $computernamebox.Text
        $querys = query session /server:"$param2" | Select-Object -skip 1 | ForEach-Object{$_.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries)}
        #write-host $querys | Format-Table
        $consoleindex = [array]::indexof($querys,"console") #logged-in user satırı console ile başladığından 
        $user = $querys[$consoleindex+1]
        $sessionid = $querys[$consoleindex+2]
        $usernamebox.Text = $user
        if ($null -eq $user) {
            $logbox_date = (get-date -format "dd.MM.yyyy HH:mm:ss")
            $logbox.Items.Add($logbox_date + " - " + "Kullanıcı kodu elde edilemedi")
        } else {
            $logbox_date = (get-date -format "dd.MM.yyyy HH:mm:ss")
            $logbox.Items.Add($logbox_date + " - " + $user + " kullanıcı kodu elde edildi")
                }
        $user; $sessionid # $asd[0]=user $asd[1]=sessionid
    }

#$ApplicationForm.Add_Shown({$ApplicationForm.Activate()}) ne iş yapar?

$form.ShowDialog()

#remove-item -path function:start*
#bir değişkeni sildin diyelim, hafızada kaldığı için program çalışmaya devam edebiliyor .s
