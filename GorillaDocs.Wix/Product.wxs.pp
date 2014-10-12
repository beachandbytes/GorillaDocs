<?xml version="1.0" encoding="UTF-8"?>
<?include Config.wxi ?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" Name="$(var.Property_ProductName)" Language="1033" Version="$(var.Property_ProductVersion)" Manufacturer="MacroView" UpgradeCode="29BEC26E-3DC0-4978-97AE-2010F841B2CC">
    <Package InstallerVersion="300" Comments="Version: $(var.Property_ProductVersion)" Compressed="yes" InstallPrivileges="elevated"/>
    <Media Id="1" Cabinet="setup.cab" EmbedCab="yes" />

    <Condition Message="[AdminMessage]">Privileged</Condition>
    <Condition Message="A later version of [ProductName] is already installed.">NOT NEWERVERSIONDETECTED</Condition>
    <Condition Message="[ProductName] requires Microsoft Visual Studio Tools for Office 4.0. Please install the VSTO 4.0 and run this installer again.">VSTO4REQUIRED &lt;&gt; 1</Condition>
    <PropertyRef Id="NETFRAMEWORK40FULL"/>
    <Condition Message="[ProductName] requires Microsoft .NET Framework 4.0 Full Profile. Please install the .NET Framework and run this installer again.">
      <![CDATA[Installed OR NETFRAMEWORK40FULL]]>
    </Condition>

    <Upgrade Id="29BEC26E-3DC0-4978-97AE-2010F841B2CC">
      <UpgradeVersion Minimum="$(var.Property_ProductVersion)" Property="NEWERVERSIONDETECTED" OnlyDetect="yes" IncludeMinimum="yes" />
      <UpgradeVersion Minimum="0.0.0.0" Maximum="$(var.Property_ProductVersion)" Property="OLDERVERSIONBEINGUPGRADED" IncludeMinimum="yes" />
    </Upgrade>

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="$(var.Property_ProgramFilesFolder)">

        <Directory Id="$(var.Property_CommonFilesFolder)">
          <Directory Id="OUTLOOKSECURITYMANAGERLOCATION" Name="Outlook Security Manager">
            <Component Id="OutlookSecurityManagerFolder" Guid="1FE86B1A-FE86-496A-A8C5-930DDFE27AFE" KeyPath="yes">
              <CreateFolder Directory="OUTLOOKSECURITYMANAGERLOCATION"/>
            </Component>
            <Component Id="SecManDlls" Guid="A79B88FC-2724-4D34-9166-CD9AE82D2571">
              <File Id="secman.dll" Name="secman.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\Aspose\secman.dll" />
              <File Id="secman64.dll" Name="secman64.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\Aspose\secman64.dll" />
            </Component>
          </Directory>
        </Directory>

        <Directory Id="INSTALLLOCATION" Name="MacroView">
          <Directory Id="PRODUCTLOCATION" Name="$(var.Property_ProductName)">

            <Component Id="ProductLocationComponent" Guid="8BBF142B-1A75-481D-808F-6149B1208805" KeyPath="yes">
              <CreateFolder Directory="PRODUCTLOCATION" />
            </Component>
            <Component Id="MacroView.Office.dll" Guid="4FBD7F50-70D7-40F5-B7B9-8EC59C17ADE5">
              <File Id="MacroView.Office.dll" Name="MacroView.Office.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\MacroView.Office\bin\MacroView.Office.dll" />
              <File Id="MacroView.Office.dll.config" Name="MacroView.Office.dll.config" Vital="yes" DiskId="1" Source="..\MacroView.Office\bin\MacroView.Office.dll.config" />
            </Component>
            <Component Id="MacroView.Word.Common.dll" Guid="EAAA491F-94D6-47BB-939F-AABD38292DED">
              <File Id="MacroView.Word.Common.dll" Name="MacroView.Word.Common.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\MacroView.Word.Common\bin\MacroView.Word.Common.dll" />
            </Component>
            <Component Id="log4net" Guid="DE0016F4-D046-4258-AD6B-B0021AD15A6D" KeyPath="yes" Win64="$(var.Property_Win64)">
              <File Id="log4net.dll" Name="log4net.dll" Source="..\packages\log4net.2.0.3\lib\net40-full\log4net.dll" Vital="yes" DiskId="1" />
            </Component>
            <Component Id="PostSharp" Guid="DE4CF0B9-53C5-4627-B531-0DFF5CF68A26">
              <File Id="PostSharp.dll" Name="PostSharp.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\PostSharp.3.1.49\lib\net20\PostSharp.dll" />
            </Component>
            <Component Id="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" Guid="32133F6D-0D2A-4998-9C31-B7BFEA526F15">
              <File Id="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" Name="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="C:\Program Files\Reference Assemblies\Microsoft\VSTO40\v4.0.Framework\Microsoft.Office.Tools.Common.v4.0.Utilities.dll" />
            </Component>
            <Component Id="GorillaDocs" Guid="6EB66AB7-78B5-4F70-BAB5-45F953575ABF">
              <File Id="GorillaDocs.dll" Name="GorillaDocs.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\lib\net40\GorillaDocs.dll" />
              <File Id="GorillaDocs.Word.dll" Name="GorillaDocs.Word.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.Word.1.0.0.0\lib\net40\GorillaDocs.Word.dll" />
              <File Id="GorillaDocs.SharePoint.dll" Name="GorillaDocs.SharePoint.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.SharePoint.1.0.0.0\lib\net40\GorillaDocs.SharePoint.dll" />
              <File Id="Microsoft.SharePoint.Client.dll" Name="Microsoft.SharePoint.Client.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.SharePoint.1.0.0.0\lib\net40\Microsoft.SharePoint.Client.dll" />
              <File Id="Microsoft.SharePoint.Client.Runtime.dll" Name="Microsoft.SharePoint.Client.Runtime.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.SharePoint.1.0.0.0\lib\net40\Microsoft.SharePoint.Client.Runtime.dll" />
              <File Id="DocumentFormat.OpenXml.dll" Name="DocumentFormat.OpenXml.dll" Vital="yes" DiskId="1" Source="..\packages\DocumentFormat.OpenXml.2.5\lib\DocumentFormat.OpenXml.dll" />
              <File Id="SecurityManager.2005.dll" Name="SecurityManager.2005.dll" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\lib\net40\SecurityManager.2005.dll" />
            </Component>
            <Component Id="Fluent" Guid="273218CB-57C1-4498-9FD5-1DFF3C5D5B29">
              <File Id="Fluent.dll" Name="Fluent.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\Fluent.Ribbon.2.0.2\lib\net40\Fluent.dll" />
              <File Id="System.Windows.Interactivity.dll" Name="System.Windows.Interactivity.dll" Vital="yes" DiskId="1" Source="..\packages\Fluent.Ribbon.2.0.2\lib\net40\System.Windows.Interactivity.dll" />
            </Component>

            <Directory Id="ChineseSimplified" Name="zh-CHS">
              <Component Id="GorillaDocs.LanguageCHS" Guid="CF76688A-E8C3-49EB-9D3D-336E600CC9E1">
                <File Id="GorillaDocs.resources.dll.chs" Name="GorillaDocs.resources.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\lib\net40\zh-CHS\GorillaDocs.resources.dll" />
                <File Id="MacroView.Office.resources.dll.chs" Name="MacroView.Office.resources.dll" Vital="yes" DiskId="1" Source="..\MacroView.Office\bin\zh-CHS\MacroView.Office.resources.dll" />

              </Component>
            </Directory>

            <Directory Id="ChineseTraditional" Name="zh-CHT">
              <Component Id="GorillaDocs.LanguageCHT" Guid="3827C27F-4847-45A7-B13B-0363072EAFFA">
                <File Id="GorillaDocs.resources.dll.cht" Name="GorillaDocs.resources.dll" KeyPath="yes" Vital="yes" DiskId="1" Source="..\packages\GorillaDocs.1.0.0.0\lib\net40\zh-CHT\GorillaDocs.resources.dll" />
              </Component>
            </Directory>

            <Directory Id="OfficeFilesDir" Name="Office Files">
              <Directory Id="ElementsDir" Name="Elements">
                <Component Id="Elements" Guid="CA233BF4-603C-41BD-8AA1-156A22F07B51">
                  <File Id="MacroViewLogo.jpg" Name="MacroViewLogo.jpg" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\MacroViewLogo.jpg" />
                  <File Id="Callouts.docx" Name="Callouts.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Callouts.docx" />
                  <File Id="MacroView.thmx" Name="MacroView.thmx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\MacroView.thmx" />
                  <File Id="Styles.dotx" Name="Styles.dotx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Styles.dotx" />
                </Component>

                <Directory Id="HeaderFooterDir" Name="HeaderFooter">
                  <Component Id="HeaderFooters" Guid="3498F0BF-F343-435E-ABA9-6CD80BF789BE">
                    <File Id="FooterPartnerLogo.docx" Name="Footer - Partner logo.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\HeaderFooter\Footer - Partner logo.docx" />
                    <File Id="FooterWebAddress.docx" Name="Footer - Web address.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\HeaderFooter\Footer - Web address.docx" />
                    <File Id="HeaderBlueLine.docx" Name="Header - Blue line.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\HeaderFooter\Header - Blue line.docx" />
                    <File Id="HeaderFirmAddress.docx" Name="Header - Firm address.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\HeaderFooter\Header - Firm address.docx" />
                  </Component>
                </Directory>

                <Directory Id="MockDataDir" Name="MockData">
                  <Component Id="MockData" Guid="7B11CC12-7C46-44C1-A039-40CDC6EB0360">
                    <File Id="LetterDetails.xml" Name="LetterDetails.xml" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\MockData\LetterDetails.xml" />
                    <File Id="AgreementDetails.xml" Name="AgreementDetails.xml" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\MockData\AgreementDetails.xml" />
                    <File Id="DocumentBuilderDetails.xml" Name="DocumentBuilderDetails.xml" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\MockData\DocumentBuilderDetails.xml" />
                  </Component>
                </Directory>

                <Directory Id="SectionsDir" Name="Sections">
                  <Component Id="Sections" Guid="1AD0C18E-404A-416C-959D-CF030A4A77C0">
                    <File Id="AgreementBody.docx" Name="Agreement Body.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Agreement Body.docx" />
                    <File Id="AgreementFrontCover.docx" Name="Agreement Front Cover.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Agreement Front Cover.docx" />
                    <File Id="AgreementTableofContents.docx" Name="Agreement Table of Contents.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Agreement Table of Contents.docx" />
                    <File Id="Annexure.docx" Name="Annexure.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Annexure.docx" />
                    <File Id="Appendix.docx" Name="Appendix.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Appendix.docx" />
                    <File Id="BackCover.docx" Name="Back Cover.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Back Cover.docx" />
                    <File Id="BlueDivider.docx" Name="Blue Divider.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Blue Divider.docx" />
                    <File Id="ExecutiveSummary.docx" Name="Executive Summary.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Executive Summary.docx" />
                    <File Id="Exhibit.docx" Name="Exhibit.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Exhibit.docx" />
                    <File Id="FrontCover.docx" Name="Front Cover.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Front Cover.docx" />
                    <File Id="Landscape.docx" Name="Landscape.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Landscape.docx" />
                    <File Id="Portrait.docx" Name="Portrait.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Portrait.docx" />
                    <File Id="Schedule.docx" Name="Schedule.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Schedule.docx" />
                    <File Id="TableofContents.docx" Name="Table of Contents.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\Table of Contents.docx" />
                    <File Id="WhiteDivider.docx" Name="White Divider.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Sections\White Divider.docx" />
                  </Component>
                </Directory>
                <Directory Id="SignoffsDir" Name="Signoffs">
                  <Component Id="Signoffs" Guid="49F4BF1E-F6D4-4DA3-B66B-87C745977DB3">
                    <File Id="SignoffDouble.docx" Name="Signoff Double.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Signoffs\Signoff Double.docx" />
                    <File Id="SignoffLHS.docx" Name="Signoff LHS.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Signoffs\Signoff LHS.docx" />
                    <File Id="SignoffRHS.docx" Name="Signoff RHS.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Elements\Signoffs\Signoff RHS.docx" />
                  </Component>
                </Directory>

              </Directory>
              <Directory Id="TemplatesDir" Name="Templates">
                <Directory Id="AgreementsDir" Name="Agreements">
                  <Component Id="Agreements" Guid="9B529092-48FA-4B95-A2FE-36F27C60144B">
                    <File Id="Agreement.dotx" Name="Agreement.dotx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Agreements\Agreement.dotx" />
                  </Component>
                </Directory>
                <Directory Id="CorrespondenceDir" Name="Correspondence">
                  <Component Id="Correspondence" Guid="AB854094-AC4A-4557-B35A-D767F4EA9A98">
                    <File Id="Letter.dotx" Name="Letter.dotx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Correspondence\Letter.dotx" />
                    <File Id="Blank.docx" Name="Blank Document.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Correspondence\Blank Document.docx" />
                  </Component>
                </Directory>
                <Directory Id="BusinessDevelopmentDir" Name="Business Development">
                  <Component Id="BusinessDevelopment" Guid="BFAC88F3-E6E0-49B6-897E-59ACC68367FB">
                    <!--<File Id="Quotation.docx" Name="Quotation.docx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Business Development\Quotation.docx" />-->
                    <File Id="DocumentBuilder.dotx" Name="Document Builder.dotx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Business Development\Document Builder.dotx" />
                  </Component>
                </Directory>
                <!--<Directory Id="PoliciesAndProceduresDir" Name="Policies and Procedures">
                  <Component Id="PoliciesAndProcedures" Guid="EA49CCF7-120F-4F4C-B7F2-1A03FC1C887C">
                    <File Id="LeaveRequest.dotx" Name="LeaveRequest.dotx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Policies and Procedures\Leave Request.dotx" />
                  </Component>
                </Directory>-->
                <Directory Id="PresentationsDir" Name="Presentations">
                  <Component Id="Presentations" Guid="86A61FFB-7DF7-4CD4-9332-C7FD8625FE7D">
                    <File Id="MacroViewPresentation.potx" Name="MacroView Presentation.potx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Presentations\MacroView Presentation.potx" />
                    <File Id="BlankPresentation.pptx" Name="Blank Presentation.pptx" Vital="yes" DiskId="1" Source="..\MacroView.Office\Office Files\Templates\Presentations\Blank Presentation.pptx" />
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
      <ComponentRef Id="HeaderFooters"/>
      <ComponentRef Id="MockData"/>
      <ComponentRef Id="Sections"/>
      <ComponentRef Id="Signoffs"/>
      <ComponentRef Id="Agreements"/>
      <ComponentRef Id="Correspondence"/>
      <ComponentRef Id="BusinessDevelopment"/>
      <ComponentRef Id="Presentations"/>
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
