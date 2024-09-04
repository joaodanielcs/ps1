Write-Host -ForegroundColor Green "Baixar e Instalar Office2021...."

$exitcode = winget uninstall --id "ProPlus2021Volume - pt-br" --disable-interactivity --accept-source-agreements
echo '<Configuration ID="cae3f462-2948-4e92-92f9-cd428499f157">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
      <Language ID="pt-br" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Publisher" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <Remove All="TRUE" />
  <RemoveMSI>
    <IgnoreProduct ID="PrjPro" />
    <IgnoreProduct ID="PrjStd" />
    <IgnoreProduct ID="VisPro" />
    <IgnoreProduct ID="VisStd" />
  </RemoveMSI>
  <AppSettings>
    <User Key="software\microsoft\office\16.0\common\general" Name="skydrivesigninoption" Value="0" Type="REG_DWORD" App="office16" Id="L_ShowSkyDriveSignIn" />
    <User Key="software\microsoft\office\16.0\common" Name="default ui theme" Value="3" Type="REG_SZ" App="office16" Id="L_DefaultUIThemeUser" />
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\excel\options" Name="disableboottoofficestart" Value="1" Type="REG_DWORD" App="excel16" Id="L_DisableOfficeStartExcel" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="disableboottoofficestart" Value="1" Type="REG_DWORD" App="ppt16" Id="L_DisableOfficeStartPowerPoint" />
    <User Key="software\microsoft\office\16.0\publisher\preferences" Name="disableboottoofficestart" Value="1" Type="REG_DWORD" App="pub16" Id="L_DisableOfficeStartPublisher" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="disableboottoofficestart" Value="1" Type="REG_DWORD" App="word16" Id="L_DisableOfficeStartWord" />
  </AppSettings>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>' > "C:\Windows\Temp\config.xml"
$exitcode = winget install "Microsoft.Office" --override "/configure C:\Windows\Temp\config.xml" --disable-interactivity --accept-source-agreements
& ([ScriptBlock]::Create((irm https://get.activated.win))) /ohook /S
