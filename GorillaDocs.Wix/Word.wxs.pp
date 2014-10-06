<?xml version="1.0" encoding="utf-8"?>
<?include Config.wxi ?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Fragment>

    <CustomAction Id="CA.WordAddin.Install" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Install" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.WordAddin.Install.SetProperty" Property="CA.WordAddin.Install" Value="OFFICEADDINKEYNAME=MacroView.Word;OFFICEVERSIONKEYNAME=[WORDVERSIONKEYNAME]" />
    <CustomAction Id="CA.WordAddin.Uninstall" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Uninstall" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.WordAddin.Uninstall.SetProperty" Property="CA.WordAddin.Uninstall" Value="OFFICEADDINKEYNAME=MacroView.Word;OFFICEVERSIONKEYNAME=[WORDVERSIONKEYNAME]" />

    <DirectoryRef Id="PRODUCTLOCATION">
      <Component Id="WordAddin" Guid="B00C68AD-2D31-4646-B5BE-14310A4EC92A" Transitive="yes" Win64="$(var.Property_Win64)">
        <Condition>(WORDVER = "Word.Application.12") OR (WORDVER = "Word.Application.14") OR (WORDVER = "Word.Application.15")</Condition>

        <File Id="MacroView.Word.dll" Name="MacroView.Word.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\MacroView.Word\bin\MacroView.Word.dll" />
        <File Id="MacroView.Word.vsto" Name="MacroView.Word.vsto" Vital="yes" DiskId="1" Source="..\MacroView.Word\bin\MacroView.Word.vsto" />
        <File Id="MacroView.Word.dll.manifest" Name="MacroView.Word.dll.manifest" Vital="yes" DiskId="1" Source="..\MacroView.Word\bin\MacroView.Word.dll.manifest" />

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[WORDVERSIONKEYNAME]\User Settings\MacroView.Word\Create\Software\Microsoft\Office\Word\Addins\MacroView Office" Root="HKLM">
          <RegistryValue Name="CommandLineSafe" Value="1" Type="integer" />
          <RegistryValue Name="LoadBehavior" Value="3" Type="integer" />
          <RegistryValue Name="FriendlyName" Value="MacroView Word Addin" Type="string" />
          <RegistryValue Name="Description" Value="MacroView Word add-in for Word" Type="string" />
          <RegistryValue Name="Manifest" Value="file://[PRODUCTLOCATION]MacroView.Word.vsto|vstolocal" Type="string" />
        </RegistryKey>

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[WORDVERSIONKEYNAME]\User Settings\MacroView.Word\Create\Software\Microsoft\VSTO\Security\Inclusion\5BB0C20E-37E0-4062-822E-79EE5737BAD7" Root="HKLM">
          <RegistryValue Name="PublicKey" Value="[MacroViewPublicKey]" Type="string" />
          <RegistryValue Name="Url" Value="file://[#MacroView.Word.vsto]" Type="string" />
        </RegistryKey>

      </Component>

    </DirectoryRef>
  </Fragment>
</Wix>