#
#                       CREATED BY ANDREW TANNER IN OCTOBER 2018
#                                     VERSION 1.0
#
PowerShell.exe -windowstyle hidden { 
#Generated Form Function
function GenerateForm {

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#Initialize Form Objects
$HelpDeskForm = New-Object System.Windows.Forms.Form
$CallStoreButton = New-Object System.Windows.Forms.Button
$GetNumButton = New-Object System.Windows.Forms.Button
$EnterLabel = New-Object System.Windows.Forms.Label
$madeBy = New-Object System.Windows.Forms.Label
$textBox1 = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

$path = 'C:\Users\' + $env:UserName + '\AppData\vanmar.mdb'

#Custom Keylistener for Enter to run the button event handler
$textbox1_KeyDown = [System.Windows.Forms.KeyEventHandler]{
    if ($_.KeyCode -eq 'Enter')
    {
        $handler_GetNumButton_Click.Invoke()
        #Suppress sound from unexpected use of enter on keyPress/keyUp
        $_.SuppressKeyPress = $true
    }
}
$textbox1.add_KeyDown($textbox1_KeyDown)

#Event Handler that gets store number and copies to clipboard
$handler_CallStoreButton_Click= 
{
    $text = $textBox1.Text
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $objConnection = New-Object -comobject ADODB.Connection
    $objRecordset = New-Object -comobject ADODB.Recordset
    $objConnection.Open("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = $path")
    $objRecordset.Open("Select * from StoreData", $objConnection,$adOpenStatic,$adLockOptimistic)
    $objRecordset.MoveFirst()
    do {$objRecordset.Fields.Item("PhoneNumber").Value; if($objRecordset.Fields.Item("StoreID").Value -eq $text){$number = $objRecordset.Fields.Item("PhoneNumber").Value;} $objRecordset.MoveNext()} until

        ($objRecordset.EOF -eq $True)

    $objRecordset.Close()
    $objConnection.Close()
    Set-Clipboard $number
    $IE= new-object -ComObject "InternetExplorer.Application"
    $IE.navigate2(“CISCOTEL://" + $number)
}
$handler_GetNumButton_Click=
{
    $text = $textBox1.Text
    $adOpenStatic = 3
    $adLockOptimistic = 3
    $objConnection = New-Object -comobject ADODB.Connection
    $objRecordset = New-Object -comobject ADODB.Recordset
    $objConnection.Open("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Users\axt3962\AppData\vanmar.mdb")
    $objRecordset.Open("Select * from StoreData", $objConnection,$adOpenStatic,$adLockOptimistic)
    $objRecordset.MoveFirst()
    do {$objRecordset.Fields.Item("PhoneNumber").Value; if($objRecordset.Fields.Item("StoreID").Value -eq $text){$number = $objRecordset.Fields.Item("PhoneNumber").Value;} $objRecordset.MoveNext()} until

    ($objRecordset.EOF -eq $True)

    $objRecordset.Close()
    $objConnection.Close()
    Set-Clipboard $number
}
$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
$HelpDeskForm.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#Form Style Code
$HelpDeskForm.Text = "Phone Wizard"
$HelpDeskForm.Name = "HelpDeskForm"
$HelpDeskForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('Y:\AndrewT\wizardcolonel.ico')
$HelpDeskForm.ShowIcon = [System.Drawing.Icon]::ExtractAssociatedIcon('Y:\AndrewT\wizardcolonel.ico')
$HelpDeskForm.DataBindings.DefaultDataSourceUpdateMode = 0
$HelpDeskForm.TopMost = $true
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 250
$System_Drawing_Size.Height = 80
$HelpDeskForm.ClientSize = $System_Drawing_Size

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 240
$System_Drawing_Size.Height = 23

#Number Button Attributes
$CallStoreButton.TabIndex = 0
$CallStoreButton.Name = "Call Store"
$CallStoreButton.UseVisualStyleBackColor = $True
$CallStoreButton.Text = "Call Store"
$CallStoreButton.Location = '5,57'
$CallStoreButton.Size = '240,20'
$CallStoreButton.DataBindings.DefaultDataSourceUpdateMode = 0
$CallStoreButton.add_Click($handler_CallStoreButton_Click)
$CallStoreButton.BackColor = 'darkred'
$CallStoreButton.ForeColor = 'white'
$CallStoreButton.Font = 'Segoe UI, style=bold'

$GetNumButton.Name = "GetNum"
$GetNumButton.add_Click($handler_GetNumButton_Click)

#TextBox Attributes
$textBox1.Location = '5,30'
$textBox1.Size = '240,20'
$textBox1.Font = 'Segoe UI, style=bold'
$textBox1.TextAlign = 'Center'
$textbox1.MaxLength = '7'

#EnterLable Attributes
$EnterLabel.Text = "Enter the Store ID:"
$EnterLabel.Font = 'Segoe UI, style=bold'
$EnterLabel.Location = '5,5'
$EnterLabel.Size = '240,40'

#madeBy Attributes
$madeBy.Text = "Made by Andrew Tanner 2018"
$madeBy.Font = 'Segoe UI'
$madeBy.Location = '5,81'
$madeBy.Size = '240,40'

#Add the elements to the window
$HelpDeskForm.Controls.Add($CallStoreButton)
$HelpDeskForm.Controls.Add($textBox1)
$HelpDeskForm.Controls.Add($EnterLabel)
$HelpDeskForm.Controls.Add($madeBy)

#endregion Form Code

#Save the initial state of the form
$InitialFormWindowState = $HelpDeskForm.WindowState
#Init the OnLoad event to correct the initial state of the form
$HelpDeskForm.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$HelpDeskForm.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm 
}