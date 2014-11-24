<?xml version="1.0" encoding="UTF-8"?>
<?include Config.wxi ?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" Name="$(var.Property_ProductName)" Language="1033" Version="$(var.Property_ProductVersion)" Manufacturer="MacroView" UpgradeCode="$(var.Property_UpgradeGUID)">
    <Package InstallerVersion="300" Comments="Version: $(var.Property_ProductVersion)" Compressed="yes" InstallPrivileges="elevated"/>
    <Media Id="1" Cabinet="setup.cab" EmbedCab="yes" />

    <Condition Message="[AdminMessage]">Privileged</Condition>
    <Condition Message="A later version of [ProductName] is already installed.">NOT NEWERVERSIONDETECTED</Condition>
    <Condition Message="[ProductName] requires Microsoft Visual Studio Tools for Office 4.0. Please install the VSTO 4.0 and run this installer again.">VSTO4REQUIRED &lt;&gt; 1</Condition>
    <PropertyRef Id="NETFRAMEWORK40FULL"/>
    <Condition Message="[ProductName] requires Microsoft .NET Framework 4.0 Full Profile. Please install the .NET Framework and run this installer again.">
      <![CDATA[Installed OR NETFRAMEWORK40FULL]]>
    </Condition>

    <Upgrade Id="$(var.Property_UpgradeGUID)">
      <UpgradeVersion Minimum="$(var.Property_ProductVersion)" Property="NEWERVERSIONDETECTED" OnlyDetect="yes" IncludeMinimum="yes" />
      <UpgradeVersion Minimum="0.0.0.0" Maximum="$(var.Property_ProductVersion)" Property="OLDERVERSIONBEINGUPGRADED" IncludeMinimum="yes" />
    </Upgrade>

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="$(var.Property_ProgramFilesFolder)">

        <Directory Id="$(var.Property_CommonFilesFolder)">
          <Directory Id="OUTLOOKSECURITYMANAGERLOCATION" Name="Outlook Security Manager">
            <Component Id="OutlookSecurityManagerFolder" Guid="[GUID]" KeyPath="yes">
              <CreateFolder Directory="OUTLOOKSECURITYMANAGERLOCATION"/>
            </Component>
            <Component Id="SecManDlls" Guid="[GUID]">
              <File Id="secman.dll" Name="secman.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\Aspose\secman.dll" />
              <File Id="secman64.dll" Name="secman64.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\Aspose\secman64.dll" />
            </Component>
          </Directory>
        </Directory>

        <Directory Id="INSTALLLOCATION" Name="MacroView">
          <Directory Id="PRODUCTLOCATION" Name="$(var.Property_ProductName)">

            <Component Id="ProductLocationComponent" Guid="[GUID]" KeyPath="yes">
              <CreateFolder Directory="PRODUCTLOCATION" />
            </Component>
            <Component Id="MacroView.Office.dll" Guid="[GUID]">
              <File Id="MacroView.Office.dll" Name="MacroView.Office.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\MacroView.Office.dll" />
              <File Id="MacroView.Office.dll.config" Name="MacroView.Office.dll.config" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\MacroView.Office.dll.config" />
            </Component>
            <Component Id="MacroView.Word.Common.dll" Guid="[GUID]">
              <File Id="MacroView.Word.Common.dll" Name="MacroView.Word.Common.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\MacroView.Word.Common.dll" />
            </Component>
            <Component Id="log4net" Guid="[GUID]" KeyPath="yes" Win64="$(var.Property_Win64)">
              <File Id="log4net.dll" Name="log4net.dll" Source="$(var.Property_WordPath)\log4net.dll" Vital="yes" DiskId="1" />
            </Component>
            <Component Id="PostSharp" Guid="[GUID]">
              <File Id="PostSharp.dll" Name="PostSharp.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\PostSharp.dll" />
            </Component>
            <Component Id="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" Guid="[GUID]">
              <File Id="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" Name="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="C:\Program Files\Reference Assemblies\Microsoft\VSTO40\v4.0.Framework\Microsoft.Office.Tools.Common.v4.0.Utilities.dll" />
            </Component>
            <Component Id="GorillaDocs" Guid="[GUID]">
              <File Id="GorillaDocs.dll" Name="GorillaDocs.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\GorillaDocs.dll" />
              <File Id="GorillaDocs.Word.dll" Name="GorillaDocs.Word.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\GorillaDocs.Word.dll" />
              <File Id="GorillaDocs.SharePoint.dll" Name="GorillaDocs.SharePoint.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\GorillaDocs.SharePoint.dll" />
              <File Id="Microsoft.SharePoint.Client.dll" Name="Microsoft.SharePoint.Client.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\Microsoft.SharePoint.Client.dll" />
              <File Id="Microsoft.SharePoint.Client.Runtime.dll" Name="Microsoft.SharePoint.Client.Runtime.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\Microsoft.SharePoint.Client.Runtime.dll" />
              <File Id="DocumentFormat.OpenXml.dll" Name="DocumentFormat.OpenXml.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\DocumentFormat.OpenXml.dll" />
              <File Id="SecurityManager.2005.dll" Name="SecurityManager.2005.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\SecurityManager.2005.dll" />
            </Component>
            <Component Id="Fluent" Guid="[GUID]">
              <File Id="Fluent.dll" Name="Fluent.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\Fluent.dll" />
              <File Id="System.Windows.Interactivity.dll" Name="System.Windows.Interactivity.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\System.Windows.Interactivity.dll" />
            </Component>

            <Directory Id="ChineseSimplified" Name="zh-CHS">
              <Component Id="GorillaDocs.LanguageCHS" Guid="[GUID]">
                <File Id="GorillaDocs.resources.dll.chs" Name="GorillaDocs.resources.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\zh-CHS\GorillaDocs.resources.dll" />
                <File Id="MacroView.Office.resources.dll.chs" Name="MacroView.Office.resources.dll" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\zh-CHS\MacroView.Office.resources.dll" />
              </Component>
            </Directory>

            <Directory Id="ChineseTraditional" Name="zh-CHT">
              <Component Id="GorillaDocs.LanguageCHT" Guid="[GUID]">
                <File Id="GorillaDocs.resources.dll.cht" Name="GorillaDocs.resources.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\zh-CHT\GorillaDocs.resources.dll" />
              </Component>
            </Directory>

            <Directory Id="OfficeFilesDir" Name="Office Files">
              <Directory Id="ElementsDir" Name="Elements">
                <Component Id="Elements" Guid="[GUID]">
                  <File Id="MacroViewLogo.jpg" Name="MacroViewLogo.jpg" Vital="yes" DiskId="1" Source="$(var.Property_AssemblyPath)\Office Files\Elements\MacroViewLogo.jpg" />
                  <File Id="Callouts.docx" Name="Callouts.docx" Vital="yes" DiskId="1" Source="$(var.Property_AssemblyPath)\Office Files\Elements\Callouts.docx" />
                  <File Id="MacroView.thmx" Name="MacroView.thmx" Vital="yes" DiskId="1" Source="$(var.Property_AssemblyPath)\Office Files\Elements\MacroView.thmx" />
                  <File Id="Styles.dotx" Name="Styles.dotx" Vital="yes" DiskId="1" Source="$(var.Property_AssemblyPath)\Office Files\Elements\Styles.dotx" />
                </Component>
              </Directory>
              <Directory Id="TemplatesDir" Name="Templates">
                <Directory Id="CorrespondenceDir" Name="Correspondence">
                  <Component Id="Correspondence" Guid="[GUID]">
                    <File Id="Letter.dotx" Name="Letter.dotx" Vital="yes" DiskId="1" Source="$(var.Property_AssemblyPath)\Office Files\Templates\Correspondence\Letter.dotx" />
                    <File Id="Blank.docx" Name="Blank Document.docx" Vital="yes" DiskId="1" Source="$(var.Property_AssemblyPath)\Office Files\Templates\Correspondence\Blank Document.docx" />
                  </Component>
                </Directory>
              </Directory>
            </Directory>

          </Directory>
        </Directory>
      </Directory>

    </Directory>

    <?if $(var.Platform) = x64 ?>
    <Binary Id="MacroView.CA" SourceFile="binary\x64\MacroView.WindowsInstaller.Actions.CA.dll" />
    <?else ?>
    <Binary Id="MacroView.CA" SourceFile="binary\x86\MacroView.WindowsInstaller.Actions.CA.dll" />
    <?endif ?>
    <!--<Icon Id="Logo" SourceFile="Binary\Generic_Document.ico" />-->

    <Property Id="DRIVE" Value="C:" />
    <Property Id="WORDVERSIONKEYNAME" Secure="yes" />
    <CustomAction Id="WORD12VERSIONKEYNAME.SetProperty" Property="WORDVERSIONKEYNAME" Value="12.0" />
    <CustomAction Id="WORD14VERSIONKEYNAME.SetProperty" Property="WORDVERSIONKEYNAME" Value="14.0" />
    <CustomAction Id="WORD15VERSIONKEYNAME.SetProperty" Property="WORDVERSIONKEYNAME" Value="15.0" />
    <Property Id="POWERPOINTVERSIONKEYNAME" Secure="yes" />
    <CustomAction Id="POWERPOINT12VERSIONKEYNAME.SetProperty" Property="POWERPOINTVERSIONKEYNAME" Value="12.0" />
    <CustomAction Id="POWERPOINT14VERSIONKEYNAME.SetProperty" Property="POWERPOINTVERSIONKEYNAME" Value="14.0" />
    <CustomAction Id="POWERPOINT15VERSIONKEYNAME.SetProperty" Property="POWERPOINTVERSIONKEYNAME" Value="15.0" />

    <Property Id="ALLUSERS" Value="1" />
    <Property Id="ARPHELPLINK" Value="http://www.macroview.com.au" />
    <Property Id="ARPURLINFOABOUT" Value="http://www.macroview.com.au" />
    <Property Id="ARPURLUPDATEINFO" Value="http://www.macroview.com.au" />
    <Property Id="ARPPRODUCTICON" Value="Logo" />
    <Property Id="AdminMessage" Value="Setup requires user to be in the administrator group in order to continue the installation process. Setup is aborting as the current user is not in the administrator group." />
    <Property Id="OLDERVERSIONBEINGUPGRADED" Secure="yes" />
    <Property Id="NEWERVERSIONDETECTED" Secure="yes" />
    <Property Id="PRODUCTLOCATION" Secure="yes" />
    <Property Id="INSTALLLOCATION" Secure="yes" />
    <Property Id="TARGETDIR" Secure="yes" />
    <Property Id="USERNAME" Secure="yes" />
    <Property Id="ROOTDRIVE" Secure="yes" />
    <Property Id="MacroViewPublicKey"><![CDATA[]]></Property>

    <!--<PropertyRef Id="NETFRAMEWORK4_LEVEL"/>-->

    <!-- VSTO -->
    <Property Id="VSTO4" Secure="yes">
      <RegistrySearch Id="CHKVSTO4" Root="HKLM" Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4" Name="Install" Type="raw" Win64="no" />
    </Property>
    <SetProperty Id="VSTO4REQUIRED" Value="1" Before="LaunchConditions">(VSTO4="" AND (WORDVER = "Word.Application.12" ) AND (POWERPOINTVER = "PowerPoint.Application.12" ))</SetProperty>

    <Property Id="WORDVER" Secure="yes">
      <RegistrySearch Id="CHKWORDVER" Root="HKLM" Key="SOFTWARE\Classes\Word.Application\CurVer" Type="raw" Win64="$(var.Property_Win64)"/>
    </Property>

    <Property Id="POWERPOINTVER" Secure="yes">
      <RegistrySearch Id="CHKPOWERPOINTVER" Root="HKLM" Key="SOFTWARE\Classes\PowerPoint.Application\CurVer" Type="raw" Win64="$(var.Property_Win64)"/>
    </Property>

    <Property Id="OFFICE2007PIAINSTALLED">
      <ComponentSearch Id="Office2007PIASearch" Guid="0638C49D-BB8B-4CD1-B191-050E8F325736" />
    </Property>
    <SetProperty Id="OFFICE2007PIAREQUIRED" Value="1" After="CostFinalize">(OFFICE2007PIAINSTALLED="" AND (WORDVER = "Word.Application.12" ) AND (POWERPOINTVER = "PowerPoint.Application.12" ))</SetProperty>

    <Feature Id="CoreFeature" Title="MacroView Office Addin" Level="1" Absent="disallow" Display="expand" AllowAdvertise="no">
      <ComponentRef Id="ProductLocationComponent" />
      <ComponentRef Id="WordAddin"/>
      <ComponentRef Id="PowerPointAddin"/>
      <ComponentRef Id="MacroView.Office.dll"/>
      <ComponentRef Id="MacroView.Word.Common.dll"/>
      <ComponentRef Id="log4net"/>
      <ComponentRef Id="GorillaDocs"/>
      <ComponentRef Id="GorillaDocs.LanguageCHS"/>
      <ComponentRef Id="GorillaDocs.LanguageCHT"/>
      <ComponentRef Id="Fluent"/>
      <ComponentRef Id="PostSharp"/>
      <ComponentRef Id="OutlookSecurityManagerFolder"/>
      <ComponentRef Id="SecManDlls"/>
      <ComponentRef Id="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" />
      <ComponentRef Id="Elements"/>
      <ComponentRef Id="Correspondence"/>
    </Feature>

    <InstallExecuteSequence>
      <SelfRegModules Sequence="5600" />
      <SelfUnregModules Sequence="2200" />
      <RemoveExistingProducts After="InstallInitialize" />

      <!-- SetProperty Custom Actions (always run) -->
      <Custom Action="WORD12VERSIONKEYNAME.SetProperty" After="CostFinalize">WORDVER = "Word.Application.12"</Custom>
      <Custom Action="WORD14VERSIONKEYNAME.SetProperty" After="WORD12VERSIONKEYNAME.SetProperty">WORDVER = "Word.Application.14"</Custom>
      <Custom Action="WORD15VERSIONKEYNAME.SetProperty" After="WORD14VERSIONKEYNAME.SetProperty">WORDVER = "Word.Application.15"</Custom>
      <Custom Action="POWERPOINT12VERSIONKEYNAME.SetProperty" After="WORD15VERSIONKEYNAME.SetProperty">POWERPOINTVER = "PowerPoint.Application.12"</Custom>
      <Custom Action="POWERPOINT14VERSIONKEYNAME.SetProperty" After="POWERPOINT12VERSIONKEYNAME.SetProperty">POWERPOINTVER = "PowerPoint.Application.14"</Custom>
      <Custom Action="POWERPOINT15VERSIONKEYNAME.SetProperty" After="POWERPOINT14VERSIONKEYNAME.SetProperty">POWERPOINTVER = "PowerPoint.Application.15"</Custom>
      <Custom Action="CA.WordAddin.Install.SetProperty" After="POWERPOINT15VERSIONKEYNAME.SetProperty"/>
      <Custom Action="CA.WordAddin.Uninstall.SetProperty" After="CA.WordAddin.Install.SetProperty"/>
      <Custom Action="CA.PowerPointAddin.Install.SetProperty" After="CA.WordAddin.Uninstall.SetProperty"/>
      <Custom Action="CA.PowerPointAddin.Uninstall.SetProperty" After="CA.PowerPointAddin.Install.SetProperty"/>

      <!-- Install -->
      <Custom Action="CA.WordAddin.Install" After="SelfRegModules">$WordAddin&gt;2</Custom>
      <Custom Action="CA.PowerPointAddin.Install" After="CA.WordAddin.Install">$PowerPointAddin&gt;2</Custom>

      <!-- Uninstall -->
      <Custom Action="CA.WordAddin.Uninstall" After="SelfUnregModules">$WordAddin=2</Custom>
      <Custom Action="CA.PowerPointAddin.Uninstall" After="CA.WordAddin.Uninstall">$PowerPointAddin=2</Custom>

    </InstallExecuteSequence>

    <!--<UIRef Id="UI_BM"/>-->

  </Product>
</Wix>
