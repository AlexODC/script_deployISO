Add-Type -AssemblyName PresentationFramework

##################################################################################################
#Déclaration des paramétres d'installation d'application.

<# 
Pour ajouter une application à la liste de téléchargement
    Ajouter dans la liste $app_Dispo le nom de l'application
    Ajouter dans la liste $app_URL_Downloader l'URL de download
    Ajouter dans la liste $app_Package le nom de l'executable

    Pour le reste c'est logique :))
#>
$app_Dispo = @('1. Installer Office',
    '2. Installer OwnCloud (Attention REBOOT Automatique)',
    '3. Installer CyberReason'
)

$app_URL_Downloader = @('',
    'https://o360.odc.fr/s/b2RDcDbQAOw2bMQ/download',
    'https://o360.odc.fr/s/k7OT8FI8UXYdeYG/download'
)

$app_Package = @('',
    'ownCloud-5.2.1.13040.x64.msi',
    'CybereasonSensor.exe'
)

##################################################################################################
<# 
Pour ajouter une version d'office à la liste des téléchargements 
    Ajouter dans la liste $office_dispo le nom de la version office, ATTENTION NE PAS TOUCHER AU 6 (Il est dans le code en dur)
    Ajouter dans la liste $office_product_id l'ID du produit officiel afin que le programme génère le fichier XML en automatique 

    Pour le reste c'est logique :))
#>
$office_dispo = @('1.Office 2019 Standard', 
'2.Office 2019 Pro Plus', 
'3.Office 2021 Standard', 
'4.Office 2021 Pro Plus', 
'5.Office 365 Business', 
'6.Office SPLA'
)

$office_product_id = @('Standard2019Retail', 
'ProPlus2019Retail', 
'Standard2021Retail',
'ProPlus2021Retail',
'O365BusinessRetail',
''
)

##################################################################################################
<# 

#>
$default_setup_dispo = @('1.Mettre Chrome par défaut', 
'2.Mettre Chrome dans la barre des tâches', 
'3.Désactiver le gestionnaire de mot de passe chrome', 
'4.Mettre Acrobat Reader par défaut', 
'5.Enlever microsoft store de la barre des tâches', 
'6.Enlever edge de la barre des tâches', 
'7.Installer Extension Bitwarden sur chrome', 
'8.Désactiver les notifications activation de PDF Creator',
'9.Mettre outlook dans la barre des tâches',
'10.Configuration de la boite mail'
)

##################################################################################################

$add_to_domain = "False"
$version_office = 64

$data_manger = @('BK',
'Mac DO',
'KFC',
'Pizza Pai',
'Picard',
'Sophie',
'Papy Henry',
'Mill Pate',
'Dominos',
'Rajah',
'Sandwich Carrefour',
'123 Burger',
'Pronto Pizza',
'OTacos',
'Subway',
'Royal',
'Panda Wok',
'Gourmet d Asie'
)
##################################################################################################

# Définition des couleurs
$GreenColor = "Green"
$RedColor = "Red"
$OrangeText = "DarkYellow"

# Vérifier si le script est lancé en tant qu'administrateur
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

##################################################################################################
##################################################################################################

#Déclaration de l'interface du menu principal
[xml]$XML_Menu = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Modern Script Place" Height="680" Width="430">
    <Grid Margin="0,0,10,-6">
        <Button Name="btn_maj_pc" Content="Mettre à jour le PC (Windows et applications Windows Store)" HorizontalAlignment="Center" Margin="10,10,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_rename_pc" Content="Renommer cet ordinateur et le mettre sur un domaine" HorizontalAlignment="Center" Margin="10,91,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install_software" Content="Installer des applications supplémentaires" HorizontalAlignment="Center" Margin="10,172,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_create_user" Content="Gérer les utilisateurs local" HorizontalAlignment="Center" Margin="10,253,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_default_setup" Content="Mise des applications par défaut et d'autre truc" HorizontalAlignment="Center" Margin="10,334,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="True"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Center" Margin="10,555,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_autre_outil" Content="Autre Outil" HorizontalAlignment="Center" Margin="10,415,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="True"/>
    
        <Label Name="copyright" Content="@Copyright CHOCHOIS Alex et LECOUTRE Antoine les GOAT" HorizontalAlignment="Center" Margin="0,636,0,0" VerticalAlignment="Top"/>
        <Button Name="btn_manger" Content="Manger ?" HorizontalAlignment="Center" Margin="10,670,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="True"/>
    
    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface de la mise à jour du poste
[xml]$XML_Maj_Poste = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     Title="Windows_Maj" Height="470" Width="820">
    <Grid>
        <ListBox Name="List_Command" Margin="0,0,0,90"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install_maj" Content="Installer les Maj" HorizontalAlignment="Left" Margin="400,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>

    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface d'installation des applications
[xml]$XML_Install_App = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Ultimate Installer Place" Height="470" Width="820">
    <Grid>
        <ListBox Name="List_Application" Margin="0,0,266,111"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install" Content="Installer" HorizontalAlignment="Left" Margin="593,0,0,0" VerticalAlignment="Center" Height="76" Width="181"/>
    </Grid> 
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface d'installation des offices
[xml]$XML_Install_Office = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Ultimate Office Place" Height="470" Width="820">
    <Grid>
        <ListBox Name="List_Office" Margin="0,0,266,111"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_install" Content="Installer" HorizontalAlignment="Left" Margin="593,0,0,0" VerticalAlignment="Center" Height="76" Width="181"/>
        <RadioButton GroupName="rb_32_64bit" Name="rb_32bit" Content="32 bit" HorizontalAlignment="Left" Margin="593,107,0,0" VerticalAlignment="Top"/>
        <RadioButton GroupName="rb_32_64bit" Name="rb_64bit" Content="64 bit" HorizontalAlignment="Left" Margin="593,127,0,0" VerticalAlignment="Top" IsChecked="True"/>

    </Grid> 
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface de renommage de poste
[xml]$XML_Rename_Poste = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="How to rename It (Place)" Height="470" Width="820">
    <Grid>
        <RadioButton GroupName="rb_Domaine" Name="rb_domaine_oui" Content="Oui" HorizontalAlignment="Left" Margin="10,33,0,0" VerticalAlignment="Top"/>
        <RadioButton GroupName="rb_Domaine" Name="rb_domaine_non" Content="Non" IsChecked="True" HorizontalAlignment="Left" Margin="10,53,0,0" VerticalAlignment="Top"/>
        <Label Content="Voulez-vous intégrer le PC à un domaine ?" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <Label Content="Nom du PC" HorizontalAlignment="Left" Margin="10,95,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_nom_pc" HorizontalAlignment="Left" Margin="10,121,0,0" TextWrapping="Wrap" Text="PC-" VerticalAlignment="Top" Width="234"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Content="Nom du Domaine" HorizontalAlignment="Left" Margin="400,10,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_domaine" HorizontalAlignment="Left" Margin="400,36,0,0" TextWrapping="Wrap" Text="Domaine.local" VerticalAlignment="Top" Width="234" IsEnabled="False" />
        <Button Name="btn_go" Content="Renommer le poste et/ou rejoindre le domaine" HorizontalAlignment="Left" Margin="10,267,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Content="Nom du compte de mise en domaine" HorizontalAlignment="Left" Margin="400,95,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_user" HorizontalAlignment="Left" Margin="400,121,0,0" TextWrapping="Wrap" Text="Administrateur" VerticalAlignment="Top" Width="234" IsEnabled="False" />
        <Button Name="btn_quitter_domaine" Content="Quitter le domaine" HorizontalAlignment="Left" Margin="400,267,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="True"/>
    </Grid> 
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface d'installation de création utilisateur
[xml]$XML_Create_User = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Ultimate User Manager Place" Height="470" Width="820">
    <Grid>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_create_user" Content="Creer l'utilisateur" HorizontalAlignment="Left" Margin="400,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Content="Nom d'utilisateur" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_name_user" HorizontalAlignment="Left" Margin="10,36,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250"/>
        <Label Content="Mot de passe (Oui il est en clair et alors ?!)" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_password" HorizontalAlignment="Left" Margin="10,106,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250"/>
        <Label Content="Nom d'affichage" HorizontalAlignment="Left" Margin="10,160,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_name_display" HorizontalAlignment="Left" Margin="10,185,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250"/>
        <Label Content="Description du compte (Facultatif)" HorizontalAlignment="Left" Margin="10,240,0,0" VerticalAlignment="Top"/>
        <TextBox Name="txt_description" HorizontalAlignment="Left" Margin="10,266,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="250"/>
        <ListBox Name="list_User" Margin="400,19,10,220"/>
        <Button Name="btn_delete_user" Content="Supprimer l'utilisateur" HorizontalAlignment="Left" Margin="400,266,0,0" VerticalAlignment="Top" Height="76" Width="195"/>
        <Button Name="btn_set_admin_user" Content="Set Admin User (Place)" HorizontalAlignment="Left" Margin="595,266,0,0" VerticalAlignment="Top" Height="76" Width="195"/>

    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface d'installation default setup
[xml]$XML_Default_Setup = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Ultimate Default Setup And Other (Place)" Height="470" Width="820">
    <Grid>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <ListBox Name="list_Default_Setup" Margin="0,0,266,91"/>
        <Button Name="btn_appliquer" Content="Appliquer la modification" HorizontalAlignment="Left" Margin="593,189,0,0" VerticalAlignment="Top" Height="76" Width="181"/>

    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################

#Déclaration de l'interface du menu autre
[xml]$XML_Menu_Autre = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Modern Script Place (Suite)" Height="680" Width="430">
    <Grid Margin="0,0,10,-6">
        <Button Name="btn_default_Printer" Content="Définir Imprimante par defaut" HorizontalAlignment="Center" Margin="10,10,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_compact_VHDX" Content="Compact VHDX" HorizontalAlignment="Center" Margin="10,91,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_Disable_UAC" Content="Désactiver les UAC" HorizontalAlignment="Center" Margin="10,172,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="true"/>
        <Button Name="btn_Clear_Log" Content="OK Cindy ! (Clear Log)" HorizontalAlignment="Center" Margin="10,253,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="true"/>
        <Button Name="btn_4" Content="Coming soon" HorizontalAlignment="Center" Margin="10,334,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="False"/>
        <Button Name="btn_autre_outil" Content="Autre Outil" HorizontalAlignment="Center" Margin="10,415,0,0" VerticalAlignment="Top" Height="76" Width="390" IsEnabled="False"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,555,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################


#Déclaration de l'interface du menu autre
[xml]$XML_Default_Printer = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
Title="Default Printer Place" Height="470" Width="820">
    <Grid Margin="0,0,10,-6">
        <Button Name="btn_recherche_carte_reseau" Content="Recherche Imprimante" HorizontalAlignment="Left" Margin="400,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <ListBox Name="list_carte_reseau" Margin="0,0,266,91" Grid.ColumnSpan="2"/>
        <Button Name="btn_explosion" Content="Définir par défaut" HorizontalAlignment="Left" Margin="593,189,0,0" VerticalAlignment="Top" Height="76" Width="181"/>
    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################


#Déclaration de l'interface du Compact VHDX
[xml]$XML_Compact_VHDX = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Compact VHDX Place" Height="470" Width="820">
    <Grid>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_compact_vhdx" Content="Lancement du compact VHDX" HorizontalAlignment="Left" Margin="400,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <TextBox Name="txt_repertoire_VHDX" HorizontalAlignment="Center" Margin="0,170,0,0" TextWrapping="Wrap" Text="E:\profils_users" VerticalAlignment="Top" Width="500"/>

    </Grid>
</Window>
"@

##################################################################################################
##################################################################################################


#Déclaration de l'interface Manger
[xml]$XML_Manger = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Where can i eat place" Height="470" Width="820">
    <Grid>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Button Name="btn_manger" Content="Manger !" HorizontalAlignment="Left" Margin="400,348,0,0" VerticalAlignment="Top" Height="76" Width="390"/>
        <Label Name="lbl_manger" HorizontalAlignment="Center" Margin="0,170,0,0" Content="MANGER !!!!" VerticalAlignment="Top" Width="100"/>

    </Grid>
</Window>
"@
##################################################################################################
##################################################################################################
##################################################################################################
#Setup de l'interface XML_Menu
$FormXML_Menu = (New-Object System.Xml.XmlNodeReader $XML_Menu)
$Window_Menu = [Windows.Markup.XamlReader]::Load($FormXML_Menu)

#Setup de l'interface XML_Maj_Poste
$FormXML_Maj_Poste = (New-Object System.Xml.XmlNodeReader $XML_Maj_Poste)
$Window_Maj_Poste = [Windows.Markup.XamlReader]::Load($FormXML_Maj_Poste)

#Setup de l'interface XML_Install_App
$FormXML_Install_App = (New-Object System.Xml.XmlNodeReader $XML_Install_App)
$Window_Install_App = [Windows.Markup.XamlReader]::Load($FormXML_Install_App)

#Setup de l'interface XML_Install_Office
$FormXML_Install_Office = (New-Object System.Xml.XmlNodeReader $XML_Install_Office)
$Window_Install_Office = [Windows.Markup.XamlReader]::Load($FormXML_Install_Office)

#Setup de l'interface XML_Rename_Poste
$FormXML_Rename_Poste = (New-Object System.Xml.XmlNodeReader $XML_Rename_Poste)
$Window_Rename_Poste = [Windows.Markup.XamlReader]::Load($FormXML_Rename_Poste)

#Setup de l'interface XML_Create_User
$FormXML_Create_User = (New-Object System.Xml.XmlNodeReader $XML_Create_User)
$Window_Create_User = [Windows.Markup.XamlReader]::Load($FormXML_Create_User)

#Setup de l'interface XML_Default_Setup
$FormXML_Default_Setup = (New-Object System.Xml.XmlNodeReader $XML_Default_Setup)
$Window_Default_Setup = [Windows.Markup.XamlReader]::Load($FormXML_Default_Setup)

#Setup de l'interface XML_Menu_Autre
$FormXML_Menu_Autre = (New-Object System.Xml.XmlNodeReader $XML_Menu_Autre)
$Window_Menu_Autre = [Windows.Markup.XamlReader]::Load($FormXML_Menu_Autre)

#Setup de l'interface XML_Default_Printer
$FormXML_Default_Printer = (New-Object System.Xml.XmlNodeReader $XML_Default_Printer)
$Window_Default_Printer = [Windows.Markup.XamlReader]::Load($FormXML_Default_Printer)

#Setup de l'interface XML_Menu_Autre
$FormXML_Compact_VHDX = (New-Object System.Xml.XmlNodeReader $XML_Compact_VHDX)
$Window_Compact_VHDX = [Windows.Markup.XamlReader]::Load($FormXML_Compact_VHDX)

#Setup de l'interface XML_Manger
$FormXML_Manger = (New-Object System.Xml.XmlNodeReader $XML_Manger)
$Window_Manger = [Windows.Markup.XamlReader]::Load($FormXML_Manger)

##################################################################################################
##################################################################################################
##################################################################################################

#Déclaration des actions des boutons du menu
$Window_Menu.FindName("btn_quitter").add_click({ 
    $Window_Menu.Close() 
})

$Window_Menu.FindName("btn_maj_pc").add_click({
    #$Window_Maj_Poste.ShowDialog()
    Update-PC
})

$Window_Menu.FindName("btn_rename_pc").add_click({ 
    $Window_Rename_Poste.ShowDialog()
})

$Window_Menu.FindName("btn_install_software").add_click({ 
    $Window_Install_App.ShowDialog()
})

$Window_Menu.FindName("btn_create_user").add_click({ 
    $Window_Create_User.ShowDialog()
})

$Window_Menu.FindName("btn_default_setup").add_click({ 
    $Window_Default_Setup.ShowDialog()
})

$Window_Menu.FindName("btn_autre_outil").add_click({ 
    $Window_Menu_Autre.ShowDialog()
})


$Window_Menu.FindName("btn_manger").add_click({ 
    $Window_Manger.ShowDialog()
})


##################################################################################################

#Déclaration des actions des boutons du menu MAJ
$Window_Maj_Poste.FindName("btn_quitter").add_click({ 
    $Window_Maj_Poste.Hide()
})

$Window_Maj_Poste.FindName("btn_install_maj").add_click({ 
    Update-PC
})

##################################################################################################

#Déclaration des actions des boutons du menu Rename
$Window_Rename_Poste.FindName("btn_quitter").add_click({ 
    $Window_Rename_Poste.Hide()
})

$Window_Rename_Poste.FindName("btn_quitter_domaine").add_click({ 
    Remove-Computer -UnjoinDomainCredential Domain01\Admin01 -WorkgroupName "Local"
})

$Window_Rename_Poste.FindName("rb_domaine_non").add_click({ 
    $Window_Rename_Poste.FindName("txt_domaine").IsEnabled = $false
    $Window_Rename_Poste.FindName("txt_user").IsEnabled = $false
})

$Window_Rename_Poste.FindName("rb_domaine_oui").add_click({ 
    $Window_Rename_Poste.FindName("txt_domaine").IsEnabled = $true
    $Window_Rename_Poste.FindName("txt_user").IsEnabled = $true
})

$Window_Rename_Poste.FindName("btn_go").add_click({
    $newName = $Window_Rename_Poste.FindName("txt_nom_pc").Text
    $newDomaine = $Window_Rename_Poste.FindName("txt_domaine").Text
    $compteDomaine = $Window_Rename_Poste.FindName("txt_user").Text

    if($Window_Rename_Poste.FindName("txt_user").IsEnabled){
        Add-Computer -ComputerName $newName -DomainName $newDomaine -Credential $compteDomaine
    } else {
        Rename-Computer -NewName $newName
    }
})

##################################################################################################

#Déclaration des actions des boutons du menu Install app
$Window_Install_App.FindName("btn_quitter").add_click({ 
    $Window_Install_App.Hide()
})

$Window_Install_App.FindName("btn_install").add_click({
    $app_to_install = $Window_Install_App.FindName("List_Application").selectedItems
    $app_splited = $app_to_install.split(".")[0]
    if ($app_splited = 1 ){
        $Window_Install_Office.ShowDialog()
    } else {
        Download-And-Install-App -downloadLink $app_URL_Downloader[$app_splited-1] -appName $app_Package[$app_splited-1]
    }
})

##################################################################################################

#Déclaration des actions des boutons du menu Install Office
$Window_Install_Office.FindName("btn_quitter").add_click({ 
    $Window_Install_Office.Hide()
})


$Window_Install_Office.FindName("rb_32bit").add_click({ 
    $version_office = 32
})

$Window_Install_Office.FindName("rb_64bit").add_click({ 
    $version_office = 64
})

$Window_Install_Office.FindName("btn_install").add_click({
    Write-Host "Lancement de l'installation d'office"
    $office_to_install = $Window_Install_Office.FindName("List_Office").selectedItems
    $office_splited = $office_to_install.split(".")

    $officeVersion = $office_splited[1]
    $XMLContent = 
@"
    <Configuration>
        <Add OfficeClientEdition="$version_office" Channel="Monthly">
            <Product ID="$office_product_id[$office_splited[0]-1]">
                <Language ID="fr-fr" />
            </Product>
        </Add>
        <Display Level="None" AcceptEULA="TRUE" />
        <Property Name="AUTOACTIVATE" Value="1"/>
    </Configuration>
"@
    if($office_product_id[$office_splited[0]] = 6){
        $officeVersion = "Office SPLA"
        $SPLADownloadLink = "https://o360.odc.fr/s/QfjXvKHGUnUu2Xj/download"
        $SPLAExePath = "$env:TEMP\Install_Office_2021_SPLA.exe"
            
        # Téléchargement du fichier d'installation SPLA
        Write-Host "Téléchargement de l'installation pour Office SPLA..."
        Invoke-WebRequest -Uri $SPLADownloadLink -OutFile $SPLAExePath
            
        # Exécution du fichier d'installation SPLA
        Write-Host "Installation de Office SPLA en cours..."
        Start-Process -FilePath $SPLAExePath -ArgumentList "/silent" -NoNewWindow -Wait
    } else {
        Install-Office
    }
})

##################################################################################################

#Déclaration des actions des boutons du menu Create User
$Window_Create_User.FindName("btn_quitter").add_click({ 
    $Window_Create_User.Hide()
})

$Window_Create_User.FindName("btn_create_user").add_click({
    $params = @{
        Name        = $Window_Create_User.FindName("txt_name_user").Text
        Password    = ConvertTo-SecureString $Window_Create_User.FindName("txt_password").Text -AsPlainText -Force 
        FullName    = $Window_Create_User.FindName("txt_name_display").Text
        Description = $Window_Create_User.FindName("txt_description").Text
    }
    New-LocalUser @params
    Write-Host "L'utilisateur a bien été créé !" -ForegroundColor $GreenColor
    
    $Window_Create_User.FindName("list_User").Items.Clear()
    foreach ($user in Get-LocalUser) {
        $Window_Create_User.FindName("list_User").Items.Add($user.Name + ":" + $user.Enabled)
    }
})

$Window_Create_User.FindName("btn_delete_user").add_click({ 
    $delete_user = $Window_Create_User.FindName("list_User").selectedItems
    $user_delete_splited = $delete_user.split(":")[0]
    Remove-LocalUser -Name $user_delete_splited

    Write-Host "L'utilisateur $user_splited a été supprimé !" -ForegroundColor $GreenColor
    
    $Window_Create_User.FindName("list_User").Items.Clear()
    foreach ($user in Get-LocalUser) {
        $Window_Create_User.FindName("list_User").Items.Add($user.Name + ":" + $user.Enabled)
    }
})

$Window_Create_User.FindName("btn_set_admin_user").add_click({ 

    $set_admin_user = $Window_Create_User.FindName("list_User").selectedItems
    $user_splited = $set_admin_user.split(":")[0]

    Add-LocalGroupMember -Group "Administrateurs" -Member $user_splited
    Write-Host "L'utilisateur $user_splited est administrateur du poste !" -ForegroundColor $GreenColor
})



##################################################################################################

#Déclaration des actions des boutons du menu Install app
$Window_Default_Setup.FindName("btn_quitter").add_click({ 
    $Window_Default_Setup.Hide()
})

$Window_Default_Setup.FindName("btn_appliquer").add_click({ 
    $selected_default_setup = $Window_Default_Setup.FindName("list_Default_Setup").selectedItems
    $choice_default_setup = $selected_default_setup.split(":")[0]

    Applicate-Default-Setup -choice $choice_default_setup
})

##################################################################################################

#Déclaration des actions des boutons du menu Autre
$Window_Menu_Autre.FindName("btn_default_Printer").add_click({ 
    $Window_Default_Printer.ShowDialog()
})

$Window_Menu_Autre.FindName("btn_compact_VHDX").add_click({ 
    $Window_Compact_VHDX.ShowDialog()
})

$Window_Menu_Autre.FindName("btn_Disable_UAC").add_click({ 
    Set-ItemProperty -Path REGISTRY::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0
    Write-Host "Les UAC sont a présent désactiver, merci de redémarrer le poste" -ForegroundColor $GreenColor
})

$Window_Menu_Autre.FindName("btn_Clear_Log").add_click({ 
    Ok-Cindy-Clear-Log
})




$Window_Menu_Autre.FindName("btn_quitter").add_click({ 
    $Window_Menu_Autre.Hide()
})

##################################################################################################

#Déclaration des actions des boutons du menu Imprimante par defaut
$Window_Default_Printer.FindName("btn_recherche_carte_reseau").add_click({ 
    $Window_Default_Printer.FindName("list_carte_reseau").Items.Clear()
    $ip_list = foreach ($adapter in Get-Printer) {
        $Window_Default_Printer.FindName("list_carte_reseau").Items.Add($adapter.Name)
    }
})

$Window_Default_Printer.FindName("btn_explosion").add_click({ 
    $valeur = $Window_Default_Printer.FindName("list_carte_reseau").selectedItems
    Write-Host $valeur
    $printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$valeur'"
    Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
    
    Write-Host "L'imprimante $printer est a présent par défaut !" -ForegroundColor $GreenColor
})

$Window_Default_Printer.FindName("btn_quitter").add_click({ 
    $Window_Default_Printer.Hide() 
}) 

##################################################################################################

#Déclaration des actions des boutons du menu Compact VHDX

$Window_Compact_VHDX.FindName("btn_quitter").add_click({ 
    $Window_Compact_VHDX.Hide()
})

$Window_Compact_VHDX.FindName("btn_compact_vhdx").add_click({

    $go_to_VHDX = $Window_Compact_VHDX.FindName("txt_repertoire_VHDX").Text

    if (Test-Path -Path $go_to_VHDX) {
        Write-Host "Dossier de VHDX trouvé ! " -ForegroundColor $GreenColor

        $vhdxList = Get-ChildItem $PSScriptRoot *.vhdx
        foreach($vhdx in $vhdxList){
            $diskpart = "select vdisk file=""$PSScriptRoot\$vhdx""
            attach vdisk readonly
            compact vdisk
            detach vdisk
            exit
            "
                New-Item $PSScriptRoot -Name "vhdx.txt" -ItemType "file" -Value $diskpart -Force
                diskpart /s $PSScriptRoot\vhdx.txt
                diskpart /s $PSScriptRoot\vhdx.txt
                Remove-Item "$PSScriptRoot\vhdx.txt"
        }
    } else {
        Write-Host "Aucun de dossier de VHDX trouvé dans $go_to_VHDX ! " -ForegroundColor $RedColor

    }
})


##################################################################################################

#Déclaration des actions des boutons du menu Compact VHDX

$Window_Manger.FindName("btn_quitter").add_click({ 
    $Window_Manger.Hide()
})

$Window_Manger.FindName("btn_manger").add_click({ 
    $manger = $data_manger | Get-Random
    $Window_Manger.FindName("lbl_manger").Content = $manger
})

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

##################################################################################################

Function Applicate-Default-Setup {
    Param (
        [String]$choice
    )
    Write-Host $choice
    Switch ($choice)
    {
        "1" { 
            Set-Chrome-Default
        }
        "2" { 
            Set-Chrome-Task-Bar
        }
        "3" { 
            "bloc de code (instructions)" 
        }
        "4" { 
            "bloc de code (instructions)" 
        }
        "5" { 
            "bloc de code (instructions)" 
        }
        "6" { 
            "bloc de code (instructions)" 
        }
        "7" { 
            "bloc de code (instructions)" 
        }
        "8" { 
            "bloc de code (instructions)" 
        }
        "9" { 
            "bloc de code (instructions)" 
        }
        Default { Write-Host "Aucun choix n'a était sélectionné" }
    }
}

Function Ok-Cindy-Clear-Log {
    Set-Executionpolicy RemoteSigned
    $days=2
    $IISLogPath="C:\inetpub\logs\LogFiles\"
    $ExchangeLoggingPath="C:\Program Files\Microsoft\Exchange Server\V15\Logging\"
    $ETLLoggingPath="C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\ETLTraces\"
    $ETLLoggingPath2="C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\Logs"
    Function CleanLogfiles($TargetFolder)
    {
      write-host -debug -ForegroundColor Yellow -BackgroundColor Cyan $TargetFolder

        if (Test-Path $TargetFolder) {
            $Now = Get-Date
            $LastWrite = $Now.AddDays(-$days)
            #$Files = Get-ChildItem $TargetFolder -Include *.log,*.blg, *.etl -Recurse | Where {$_.LastWriteTime -le "$LastWrite"}
           $Files = Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\"  -Recurse | Where-Object {$_.Name -like "*.log" -or $_.Name -like "*.blg" -or $_.Name -like "*.etl"}  | where {$_.lastWriteTime -le "$lastwrite"} | Select-Object FullName  
            foreach ($File in $Files)
                {
                   $FullFileName = $File.FullName  
                   Write-Host "Suppression de $FullFileName --> OK Cindy !" -ForegroundColor "yellow"; 
                    Remove-Item $FullFileName -ErrorAction SilentlyContinue | out-null
                }
           }
    Else {
        Write-Host "Le dossier $TargetFolder n'existe pas Cindy !" -ForegroundColor "red"
        }
    }
    CleanLogfiles($IISLogPath)
    CleanLogfiles($ExchangeLoggingPath)
    CleanLogfiles($ETLLoggingPath)
    CleanLogfiles($ETLLoggingPath2)
}

Function Set-Chrome-Default {
    # Définir l'ID de l'application Chrome
    $chromeAppId = "XXXXX"

    # Créer la sous-clé et la valeur
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\html" -Name "(Default)" -Value $chromeAppId -Type DWORD

    # Redémarrer l'explorateur Windows
    Restart-Process -Name "explorer"
}

Function Set-Chrome-Task-Bar {
    # Créer un objet ShellLink
    $shellLink = New-Object -ComObject Shell.Application

    # Définir la propriété TargetPath vers le fichier exécutable de Chrome
    $shellLink.TargetPath = "C:\Program Files\Google\Chrome\chrome.exe"

    # Définir la propriété Arguments pour lancer Chrome en mode maximisé
    $shellLink.Arguments = "-start-maximized"

    # Définir la propriété WorkingDirectory vers le dossier de Chrome
    $shellLink.WorkingDirectory = "C:\Program Files\Google\Chrome"

    # Définir la propriété Description
    $shellLink.Description = "Google Chrome"

    # Enregistrer le raccourci dans le dossier de la barre des tâches
    $shellLink.SaveLink($env:APPDATA + "\Microsoft\Windows\Start Menu\Programs\Start Menu\Google Chrome.lnk")

    # Épingler le raccourci à la barre des tâches
    Add-AppxPackage -Path $env:APPDATA + "\Microsoft\Windows\Start Menu\Programs\Start Menu\Google Chrome.lnk" -Register
}

# Fonction pour télécharger et installer une application depuis un lien public OwnCloud, Cybereason, etc.
Function Download-And-Install-App {
    Param (
        [string]$downloadLink,
        [string]$appName
    )
    $downloadsPath = [System.Environment]::GetFolderPath('UserProfile') + '\Downloads'
    $localPath = Join-Path -Path $downloadsPath -ChildPath $appName

    try {
        Invoke-WebRequest -Uri $downloadLink -OutFile $localPath
        if (Test-Path $localPath) {
            if ($appName -like "*.msi") {
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$localPath`" /qn" -Wait
                Write-Host "$appName a été installée avec succÃ¨s." -ForegroundColor $GreenColor
            } else {
                Start-Process -FilePath $localPath -Args "/S" -Wait
                Write-Host "$appName a été installée avec succÃ¨s." -ForegroundColor $GreenColor
            }
        } else {
            Write-Host "Erreur : Impossible de trouver le fichier téléchargé." -ForegroundColor $RedColor
        }
    } catch {
        Write-Host "Erreur lors du téléchargement ou de l'installation : $_" -ForegroundColor $RedColor
    }
    Remove-Item -Path $localPath -Force
}

# Fonction pour détecter la version d'Office installée
Function Detect-InstalledOffice {
    $Folder = 'C:\Program Files\Microsoft Office 15'
    if (Test-Path -Path $Folder) {
        Write-Host "Un office est déja présent !"
        return true
    } else {
        return false
    }
}

# Fonction pour installer Office
Function Install-Office {
    $ODTDownloadLink = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17126-20132.exe"
    $ODTPath = "$env:TEMP\ODT"
    $XMLPath = "$ODTPath\configuration.xml"

    # Téléchargement et extraction de l'Office Deployment Tool
    Invoke-WebRequest -Uri $ODTDownloadLink -OutFile "$env:TEMP\ODT.exe"
    Start-Process -FilePath "$env:TEMP\ODT.exe" -ArgumentList "/extract:`"$ODTPath`"" -NoNewWindow -Wait

    $XMLContent | Out-File -FilePath $XMLPath

    # Détecter et désinstaller la version actuelle d'Office si présente
    $installedOffice = Detect-InstalledOffice
    if ($installedOffice) {
        Write-Host "Désinstallation de la version actuelle d'Office..."
        Remove-PreviousOfficeInstalls -Force -Quiet
    }

    # Installation de la nouvelle version d'Office
    Write-Host "Installation de $officeVersion en cours..." -ForegroundColor $GreenColor
    Start-Process -FilePath "$ODTPath\setup.exe" -ArgumentList "/configure `"$XMLPath`"" -NoNewWindow -Wait
}

# Fonction pour installer toutes les mises à jour Windows, y compris les facultatives
Function Install-WindowsUpdates {
    Write-Host "Lancement des mises à jours Windows"
    Check-And-Install-NuGet

    if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Remove-Module -Name PSWindowsUpdate -Force -ErrorAction SilentlyContinue
    }
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber

    $updates = Get-WindowsUpdate -AcceptAll -IgnoreReboot
    $rebootRequired = $false

    if ($updates.Count -gt 0) {
        foreach ($update in $updates) {
            Write-Host "Installation de la mise à jour : $($update.Title)" -ForegroundColor $GreenColor
            Install-WindowsUpdate -KBArticleID $update.KBArticleID -AutoReboot:$false -Confirm:$false -IgnoreRebootRequired
            if ($update.IsRebootRequired) {
                $rebootRequired = $true
            }
        }
        if ($rebootRequired) {
            Write-Host "Redémarrage nécessaire pour terminer l'installation des mises à jour. Veuillez redémarrer votre ordinateur." -ForegroundColor $OrangeText
        } else {
            Write-Host "Toutes les mises à jour ont été installallées. Un redémarrage pourrait être nécessaire pour appliquer les mises à jour." -ForegroundColor $OrangeText
        }
    } else {
        Write-Host "Aucune mise à jour Windows disponible." -ForegroundColor $GreenColor
    }
    return
}

# Fonction pour mettre à jour les applications du Windows Store
Function Update-WindowsStoreApps {
    Write-Host "Lancement des mises à jours des applications" -ForegroundColor $GreenColor
    $wingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    if (Test-Path $wingetPath) {
        winget upgrade --all --force --accept-package-agreements --accept-source-agreements
        Write-Host "Les mises à jour des applications Windows Store ont été effectuées." -ForegroundColor $GreenColor
    } else {
        Write-Host "Winget n'est pas installé." -ForegroundColor $RedColor
    }
    return 
}

# Fonction pour mettre à jour le PC (Windows et applications Windows Store)
Function Update-PC {
    Install-WindowsUpdates
    Update-WindowsStoreApps
}

##################################################################################################
# Mise en place des listes d'applications et d'office dans l'interface graphique
foreach ($office in $office_dispo) {
    $Window_Install_Office.FindName("List_Office").Items.Add("$office")
}
foreach ($app in $app_Dispo) {
    $Window_Install_App.FindName("List_Application").Items.Add("$app")
}
foreach ($user in Get-LocalUser) {
    $Window_Create_User.FindName("list_User").Items.Add($user.Name + ":" + $user.Enabled)
}
foreach ($setup_dispo in $default_setup_dispo) {
    $Window_Default_Setup.FindName("list_Default_Setup").Items.Add("$setup_dispo")
}
foreach ($adapter in Get-Printer) {
    $Window_Default_Printer.FindName("list_carte_reseau").Items.Add($adapter.Name)
}

##################################################################################################

cls
Write-Host "Boite de dialogue permettant la visualisation des commandes, ne pas fermer cette fenêtre."
#Affichage de la interface et début du programme !
$Window_Menu.ShowDialog()