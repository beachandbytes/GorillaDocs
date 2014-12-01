<?xml version="1.0" encoding="utf-8"?>
<?include Config.wxi ?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Fragment>

    <CustomAction Id="CA.WordAddin.Install" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Install" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.WordAddin.Install.SetProperty" Property="CA.WordAddin.Install" Value="OFFICEADDINKEYNAME=$(var.Property_WordAssemblyName);OFFICEVERSIONKEYNAME=[WORDVERSIONKEYNAME]" />
    <CustomAction Id="CA.WordAddin.Uninstall" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Uninstall" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.WordAddin.Uninstall.SetProperty" Property="CA.WordAddin.Uninstall" Value="OFFICEADDINKEYNAME=$(var.Property_WordAssemblyName);OFFICEVERSIONKEYNAME=[WORDVERSIONKEYNAME]" />

    <DirectoryRef Id="PRODUCTLOCATION">
      <Component Id="WordAddin" Guid="[GUID]" Transitive="yes" Win64="$(var.Property_Win64)">
        <Condition>(WORDVER = "Word.Application.12") OR (WORDVER = "Word.Application.14") OR (WORDVER = "Word.Application.15")</Condition>

        <File Id="$(var.Property_WordAssemblyName).dll" Name="$(var.Property_WordAssemblyName).dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\$(var.Property_WordAssemblyName).dll" />
        <File Id="$(var.Property_WordAssemblyName).vsto" Name="$(var.Property_WordAssemblyName).vsto" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\$(var.Property_WordAssemblyName).vsto" />
        <File Id="$(var.Property_WordAssemblyName).dll.manifest" Name="$(var.Property_WordAssemblyName).dll.manifest" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\$(var.Property_WordAssemblyName).dll.manifest" />
        <File Id="$(var.Property_WordAssemblyName).dll.config" Name="$(var.Property_WordAssemblyName).dll.config" Vital="yes" DiskId="1" Source="$(var.Property_WordPath)\$(var.Property_WordAssemblyName).dll.config" />

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[WORDVERSIONKEYNAME]\User Settings\$(var.Property_WordAssemblyName)\Create\Software\Microsoft\Office\Word\Addins\$(var.Property_ProductName)" Root="HKLM">
          <RegistryValue Name="CommandLineSafe" Value="1" Type="integer" />
          <RegistryValue Name="LoadBehavior" Value="3" Type="integer" />
          <RegistryValue Name="FriendlyName" Value="$(var.Property_ClientName) Word Addin" Type="string" />
          <RegistryValue Name="Description" Value="$(var.Property_ClientName) Word add-in for Word" Type="string" />
          <RegistryValue Name="Manifest" Value="file://[PRODUCTLOCATION]$(var.Property_WordAssemblyName).vsto|vstolocal" Type="string" />
        </RegistryKey>

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[WORDVERSIONKEYNAME]\User Settings\$(var.Property_WordAssemblyName)\Create\Software\Microsoft\VSTO\Security\Inclusion\[GUID]" Root="HKLM">
          <RegistryValue Name="PublicKey" Value="[MacroViewPublicKey]" Type="string" />
          <RegistryValue Name="Url" Value="file://[#$(var.Property_WordAssemblyName).vsto]" Type="string" />
        </RegistryKey>

      </Component>

    </DirectoryRef>
  </Fragment>
</Wix>