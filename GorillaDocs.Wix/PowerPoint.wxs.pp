<?xml version="1.0" encoding="utf-8"?>
<?include Config.wxi ?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Fragment>

    <CustomAction Id="CA.PowerPointAddin.Install" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Install" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.PowerPointAddin.Install.SetProperty" Property="CA.PowerPointAddin.Install" Value="OFFICEADDINKEYNAME=$(var.Property_PowerPointAssemblyName);OFFICEVERSIONKEYNAME=[POWERPOINTVERSIONKEYNAME]" />
    <CustomAction Id="CA.PowerPointAddin.Uninstall" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Uninstall" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.PowerPointAddin.Uninstall.SetProperty" Property="CA.PowerPointAddin.Uninstall" Value="OFFICEADDINKEYNAME=$(var.Property_PowerPointAssemblyName);OFFICEVERSIONKEYNAME=[POWERPOINTVERSIONKEYNAME]" />

    <DirectoryRef Id="PRODUCTLOCATION">
      <Component Id="PowerPointAddin" Guid="[GUID]" Transitive="yes" Win64="$(var.Property_Win64)">
        <Condition>(POWERPOINTVER = "PowerPoint.Application.12") OR (POWERPOINTVER = "PowerPoint.Application.14") OR (POWERPOINTVER = "PowerPoint.Application.15")</Condition>

        <File Id="$(var.Property_PowerPointAssemblyName).dll" Name="$(var.Property_PowerPointAssemblyName).dll" KeyPath="yes" Vital="yes" DiskId="1" Source="$(var.Property_PowerPointPath)\$(var.Property_PowerPointAssemblyName).dll" />
        <File Id="$(var.Property_PowerPointAssemblyName).vsto" Name="$(var.Property_PowerPointAssemblyName).vsto" Vital="yes" DiskId="1" Source="$(var.Property_PowerPointPath)\$(var.Property_PowerPointAssemblyName).vsto" />
        <File Id="$(var.Property_PowerPointAssemblyName).dll.manifest" Name="$(var.Property_PowerPointAssemblyName).dll.manifest" Vital="yes" DiskId="1" Source="$(var.Property_PowerPointPath)\$(var.Property_PowerPointAssemblyName).dll.manifest" />
        <File Id="$(var.Property_PowerPointAssemblyName).dll.config" Name="$(var.Property_PowerPointAssemblyName).dll.config" Vital="yes" DiskId="1" Source="$(var.Property_PowerPointPath)\$(var.Property_PowerPointAssemblyName).dll.config" />

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[POWERPOINTVERSIONKEYNAME]\User Settings\$(var.Property_PowerPointAssemblyName)\Create\Software\Microsoft\Office\PowerPoint\Addins\[Property_ProductName]" Root="HKLM">
          <RegistryValue Name="CommandLineSafe" Value="1" Type="integer" />
          <RegistryValue Name="LoadBehavior" Value="3" Type="integer" />
          <RegistryValue Name="FriendlyName" Value="$(var.Property_ClientName) PowerPoint Addin" Type="string" />
          <RegistryValue Name="Description" Value="$(var.Property_ClientName) PowerPoint add-in for PowerPoint" Type="string" />
          <RegistryValue Name="Manifest" Value="file://[PRODUCTLOCATION]$(var.Property_PowerPointAssemblyName).vsto|vstolocal" Type="string" />
        </RegistryKey>

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[POWERPOINTVERSIONKEYNAME]\User Settings\$(var.Property_PowerPointAssemblyName)\Create\Software\Microsoft\VSTO\Security\Inclusion\[GUID]" Root="HKLM">
          <RegistryValue Name="PublicKey" Value="[MacroViewPublicKey]" Type="string" />
          <RegistryValue Name="Url" Value="file://[#$(var.Property_PowerPointAssemblyName).vsto]" Type="string" />
        </RegistryKey>

      </Component>

    </DirectoryRef>
  </Fragment>
</Wix>