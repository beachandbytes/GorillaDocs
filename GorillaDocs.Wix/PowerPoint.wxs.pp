<?xml version="1.0" encoding="utf-8"?>
<?include Config.wxi ?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Fragment>

    <CustomAction Id="CA.PowerPointAddin.Install" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Install" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.PowerPointAddin.Install.SetProperty" Property="CA.PowerPointAddin.Install" Value="OFFICEADDINKEYNAME=MacroView.PowerPoint;OFFICEVERSIONKEYNAME=[POWERPOINTVERSIONKEYNAME]" />
    <CustomAction Id="CA.PowerPointAddin.Uninstall" BinaryKey="MacroView.CA" DllEntry="OfficeUserSettingsRegistry_Uninstall" Execute="deferred" Impersonate="no" />
    <CustomAction Id="CA.PowerPointAddin.Uninstall.SetProperty" Property="CA.PowerPointAddin.Uninstall" Value="OFFICEADDINKEYNAME=MacroView.PowerPoint;OFFICEVERSIONKEYNAME=[POWERPOINTVERSIONKEYNAME]" />

    <DirectoryRef Id="PRODUCTLOCATION">
      <Component Id="PowerPointAddin" Guid="5D222027-D29B-48F9-AA33-1A17A041CD81" Transitive="yes" Win64="$(var.Property_Win64)">
        <Condition>(POWERPOINTVER = "PowerPoint.Application.12") OR (POWERPOINTVER = "PowerPoint.Application.14") OR (POWERPOINTVER = "PowerPoint.Application.15")</Condition>

        <File Id="MacroView.PowerPoint.dll" Name="MacroView.PowerPoint.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\MacroView.PowerPoint\bin\MacroView.PowerPoint.dll" />
        <File Id="MacroView.PowerPoint.vsto" Name="MacroView.PowerPoint.vsto" Vital="yes" DiskId="1" Source="..\MacroView.PowerPoint\bin\MacroView.PowerPoint.vsto" />
        <File Id="MacroView.PowerPoint.dll.manifest" Name="MacroView.PowerPoint.dll.manifest" Vital="yes" DiskId="1" Source="..\MacroView.PowerPoint\bin\MacroView.PowerPoint.dll.manifest" />

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[POWERPOINTVERSIONKEYNAME]\User Settings\MacroView.PowerPoint\Create\Software\Microsoft\Office\PowerPoint\Addins\MacroView Office" Root="HKLM">
          <RegistryValue Name="CommandLineSafe" Value="1" Type="integer" />
          <RegistryValue Name="LoadBehavior" Value="3" Type="integer" />
          <RegistryValue Name="FriendlyName" Value="MacroView PowerPoint Addin" Type="string" />
          <RegistryValue Name="Description" Value="MacroView PowerPoint add-in for PowerPoint" Type="string" />
          <RegistryValue Name="Manifest" Value="file://[PRODUCTLOCATION]MacroView.PowerPoint.vsto|vstolocal" Type="string" />
        </RegistryKey>

        <RegistryKey Action="createAndRemoveOnUninstall" Key="Software\Microsoft\Office\[POWERPOINTVERSIONKEYNAME]\User Settings\MacroView.PowerPoint\Create\Software\Microsoft\VSTO\Security\Inclusion\CFB38D38-8940-4E47-8596-7769FE8E94B4" Root="HKLM">
          <RegistryValue Name="PublicKey" Value="[MacroViewPublicKey]" Type="string" />
          <RegistryValue Name="Url" Value="file://[#MacroView.PowerPoint.vsto]" Type="string" />
        </RegistryKey>

      </Component>

    </DirectoryRef>
  </Fragment>
</Wix>