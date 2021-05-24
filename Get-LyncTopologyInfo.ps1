param([string]$Source = "", [switch]$vsd, [switch]$xls)

cls
##################################################
#
#  Get-LyncTopologyInfo.ps1
#
#  Author: Christopher Cook
#  Website: EmptyMessage.com
#  Version: 1.2
#  Release: 09/2011
#  
#  This work is licensed under the Creative Commons
#  Attribution 3.0 Unported License. To view a copy
#  of this license, visit
#  http://creativecommons.org/licenses/by/3.0/
#  or send a letter to
#  Creative Commons
#  444 Castro Street, Suite 900
#  Mountain View, California, 94041, USA.
#
#  Use at your own risk!
#
##################################################

function Build-Spreadsheet {
#Set the file name for the Excel Spreadsheet we're working with.
#Drop the VCFG extension and add " - Topology.XLSX" to the end.
Write-Host "Starting Excel..."
$ExcelFileName = $TBXMLFileName.substring(0,$TBXMLFileName.Length - 6) + " - Topology.XLSX"

#Create the Excel Instance we will be working with.
$ExcelApp = New-Object -comobject Excel.Application

#Create a new Workbook and add 2 extra sheets for 5 total.
Write-Host "Configuring Worksheets..."
$Workbook = $ExcelApp.Workbooks.Add()
#$Worksheets = $WorkBook.Worksheets.Add()

#Name the sheets in the Workbook and assign variables to them.
$SitesServersSheet = $Workbook.Worksheets.Item(1)
$SitesServersSheet.Name = "Lync Sites And Servers"
$ManagementInfoSheet = $Workbook.Worksheets.Item(2)
$ManagementInfoSheet.Name = "Management Info"
$ServicesSheet = $Workbook.Worksheets.Item(3)
$ServicesSheet.Name = "Services"
Write-Host "Done"


Write-Host "**Building Management Info Sheet**"

# Build SIP Domains and URL Configuration.
$CurrentRow = 1
$CurrentColumn = 1
$ManagementInfoSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Default SIP Domain:"
$ManagementInfoSheet.Cells.Item($CurrentRow + 1,$CurrentColumn) = $DefaultSIPDomain
$CurrentColumn++
foreach ($NameSpace in $SIPDomains) {
	$ManagementInfoSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Alternate SIP Domain:"
	$ManagementInfoSheet.Cells.Item($CurrentRow + 1,$CurrentColumn) = $NameSpace
	$CurrentColumn++
}
foreach ($URL in $URLS) {
	$ManagementInfoSheet.Cells.Item($CurrentRow,$CurrentColumn) = $URL.Component + " URL:"
	$ManagementInfoSheet.Cells.Item($CurrentRow + 1,$CurrentColumn) = $URL.ActiveUrl
	$CurrentColumn++
}
$ManagementInfoSheet.UsedRange.Columns.Autofit() | Out-Null
$objList = $ManagementInfoSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $ManagementInfoSheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
$objList.TableStyle = "TableStyleMedium20"

Write-Host "Done"

#Build Lync Sites and Servers worksheet.
Write-Host "**Building Sites and Servers Sheet**"
$MaxSQLInstances = 1
$MaxNetInterfaces = 1
foreach ($Server in $Servers){
if ($Server.SQLInstances.Count -gt $MaxSQLInstances){$MaxSQLInstances = $Server.SQLInstances.Count}
if ($Server.NetInterfaces.Count -gt $MaxNetInterfaces){$MaxNetInterfaces = $Server.NetInterfaces.Count}
}

$CurrentRow = 1
$CurrentColumn = 1
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Server Name:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Server Role(s):"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Single Machine:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Cluster FQDN:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Unique ID:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Number in Cluster:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Site Name:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Site Description:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Site City:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Site State:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Site Country:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Site ID:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Remote Site:"
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Parent Site:"
$CurrentColumn++

$i = 1
do {
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "SQL Instance #"+$i+":"
$CurrentColumn++
$i++
}
while ($i -le $MaxSQLInstances)

$i = 1
do {
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Network Interface #"+$i+":"
$CurrentColumn++
$i++
}
while ($i -le $MaxNetInterfaces)


foreach ($Server in $Servers){
$CurrentColumn = 1
$SiteInfo = $Sites | Where-Object{$_.OriginalSiteId -eq $Server.OriginalSiteId} | Select Name,Description,City,State,Country,OriginalSiteID,ParentSite

$CurrentRow++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Server.Name
$CurrentColumn++

$ServerServices = @()
foreach ($Service in $Services){if ($Service.InstalledOn -eq $Server.UniqueID){$ServerServices += $Service.RoleName}}
$ServerRole = @()
if ($ServerServices -contains "UserServices"){$ServerRole += "Front-End"}
if ($ServerServices -contains "FileStore"){$ServerRole += "File-Store"}
if ($ServerServices -contains "ArchivingServer"){$ServerRole += "Archiving"}
if ($ServerServices -contains "MonitoringServer"){$ServerRole += "Monitoring"}
if ($ServerServices -contains "ConfServices" -and $ServerServices -notcontains "UserServices"){$ServerRole += "AV-Conferencing"}
if ($ServerServices -contains "Registrar" -and $ServerServices -notcontains "UserServices"){$ServerRole += "Director"}
if ($ServerServices -contains "Registrar" -and !$SiteInfo.ParentSite -eq $NULL){$ServerRole += "SBA / SBS"}
if ($ServerServices -contains "MediationServer"){$ServerRole += "Mediation"}
if ($ServerServices -contains "ExternalServer"){$ServerRole += "External-App-Server"}
if ($ServerServices -contains "Registrar" -and $SiteInfo.Name -like "BackCompatSite"){$ServerRole = "OCS-Server"}
if ($ServerServices -contains "EdgeServer"){$ServerRole += "Edge"}
if ($ServerServices -contains "PSTNGateway"){$ServerRole += "PSTN-Gateway"}

if ($Server.SQLInstances){
	$Server.SQLInstances.GetEnumerator() | Foreach-Object {
		if ($_.Value.UniqueID -ne "None"){$CellData = "Name: " + $_.Value.Name + ", UniqueID: " + $_.Value.UniqueID + ", OriginalClusterUniqueID: " + $_.Value.OriginalClusterUniqueID}
		if ($ServerServices.Count -eq 0 -and $_.Value.UniqueID -like "sql*"){$ServerRole += "Back-End-SQL"}
}}


$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = " " + $ServerRole
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Server.IsSingleMachineOnly
$CurrentColumn++
if ($Server.ClusterFQDN){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Server.ClusterFQDN}
$CurrentColumn++
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = " "+$Server.UniqueID
$CurrentColumn++
if ($Server.IsSingleMachineOnly -ne "true"){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = " " + $Server.OrdinalInCluster}
$CurrentColumn++
if ($SiteInfo.Name){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.Name.ToString()}
$CurrentColumn++
if ($SiteInfo.Description){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.Description.ToString()}
$CurrentColumn++
if ($SiteInfo.City){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.City.ToString()}
$CurrentColumn++
if ($SiteInfo.State){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.State.ToString()}
$CurrentColumn++
if ($SiteInfo.Country){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.Country.ToString()}
$CurrentColumn++
if ($SiteInfo.OriginalSiteID){$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.OriginalSiteID.ToString()}
$CurrentColumn++
if ($SiteInfo.ParentSite -eq $NULL){
	$CurrentColumn++
	$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = "No";$CurrentColumn++}
else
	{$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = "Yes"
	$CurrentColumn++
	$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $SiteInfo.ParentSite.ToString()}

$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = $SQLCellData
$SQLCellData = ""
$CurrentColumn++

$Server.NetInterfaces.GetEnumerator() | Foreach-Object {
	if ($_.Value.InterfaceNumber -ne "None"){
		$CellData = "Interface Side: " + $_.Value.InterfaceSide + ", IP Address: " + $_.Value.IPAddress + ", Configured IP Address: " + $_.Value.ConfiguredIPAddress
	}
$SitesServersSheet.Cells.Item($CurrentRow,$CurrentColumn) = $CellData
$CellData = ""
$CurrentColumn++
}
$ServerServices = $Null
}
$SitesServersSheet.UsedRange.Columns.Autofit() | Out-Null
$objList = $SitesServersSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $SitesServersSheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
$objList.TableStyle = "TableStyleMedium20"
Write-Host "Done"


#Build Lync Services and Ports worksheet.
Write-Host "**Building Services and Ports Sheet**"

$CurrentRow = 1
$CurrentColumn = 1
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Role Name:"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Unique ID:"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Installed On:"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Type:"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Dependencies:"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Port Number"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Port Range"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Port Protocol"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Port Owner"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Port Usage"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Interface Side"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Interface Number"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "Port URL Path"
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = "GRUU Type"
$CurrentColumn++

foreach ($Service in $Services){

$Service.Ports.GetEnumerator() | Foreach-Object {
$arrOwner = $_.Value.Owner.Split(":")
$Owner = $arrOwner[2]

$CurrentColumn = 1
$CurrentRow++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Service.RoleName
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Service.UniqueID
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = " " + $Service.InstalledOn
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Service.Type
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Service.Dependencies
$CurrentColumn++
# Service ports.
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = " " + $_.Value.PortNumber
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = " " + $_.Value.Range
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $_.Value.Protocol
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Owner
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $_.Value.Usage
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $_.Value.InterfaceSide
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $_.Value.InterfaceNumber
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $_.Value.UrlPath
$CurrentColumn++
$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $_.Value.GruuType
$CurrentColumn++



$ServicesSheet.Cells.Item($CurrentRow,$CurrentColumn) = $CellData
$CurrentColumn++
}}

$ServicesSheet.UsedRange.Columns.Autofit() | Out-Null
$objList = $ServicesSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $ServicesSheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
$objList.TableStyle = "TableStyleMedium20"

Write-Host "Done"

Write-Host "Finished creating spreadsheet."
$ExcelApp.Visible = $True
$Workbook.SaveAs($CurrentDirectory + $ExcelFileName) | Out-Null
Write-Host "Visio Document Complete!"
}

function Build-Drawing {
#Set the file name for the Visio drawing we're working with.
#Drop the TBXML extension and add " - Topo_Map.VSD" to the end.
Write-Host "Starting Visio..."
$VisioFileName = $TBXMLFileName.substring(0,$TBXMLFileName.Length - 6) + " - Topo_Map.VSD"

#Create the Visio instance we will be working with.
$VisioApp = New-Object -ComObject Visio.Application

#Setup document, pages, and stencils we will be using.
Write-Host "Configuring documents, pages, and stencils..."
$objDocuments = $VisioApp.Documents
$objDocument = $objDocuments.Add("NETW_U.VST")
$objDocument.PaperSize = 24
$objDocument.PrintLandscape = $True
$objPages = $VisioApp.ActiveDocument.Pages
$objPage = $objPages.Item(1)
$objPage.AutoSize = $False

$StencilPath = [System.Environment]::GetFolderPath('MyDocuments') + "\My Shapes"
$StencilFileName = "LyncServer2010VisioStencil.vss"
$colLyncStencils = $VisioApp.Documents.Add($StencilPath.ToString() + "\" + $StencilFileName.ToString())
$colBasicStencils = $VisioApp.Documents.Add("BASIC_U.VSS")

$stnApplicationServer = $colLyncStencils.Masters.Item("Application Server")
$stnArchivingServer = $colLyncStencils.Masters.Item("Archiving Server")
$stnAVServer = $colLyncStencils.Masters.Item("AV Conferencing Server Role")
$stnBranchAppliance = $colLyncStencils.Masters.Item("Branch Office Appliance")
$stnCluster = $colLyncStencils.Masters.Item("Virtual Machine 2")
$stnDatabaseServer = $colLyncStencils.Masters.Item("Database Server")
$stnDirector = $colLyncStencils.Masters.Item("Director")
$stnDirectorPool = $colLyncStencils.Masters.Item("Director Array")
$stnEdgeServer = $colLyncStencils.Masters.Item("Edge Server")
$stnEdgePool = $colLyncStencils.Masters.Item("Edge Array Server")
$stnEnterpriseFE = $colLyncStencils.Masters.Item("Enterprise Edition Front End")
$stnFEPool = $colLyncStencils.Masters.Item("Enterprise Pool")
$stnFileStore = $colLyncStencils.Masters.Item("File Server")
$stnFirewall = $colLyncStencils.Masters.Item("Firewall")
$stnPBX = $colLyncStencils.Masters.Item("IP PBX")
$stnLoadBalancer = $colLyncStencils.Masters.Item("Load Balancer 3D")
$stnMediationServer = $colLyncStencils.Masters.Item("Mediation Server")
$stnMonitoringServer = $colLyncStencils.Masters.Item("Database Server")
$stnRegistrar = $colLyncStencils.Masters.Item("Registrar")
$stnReverseProxy = $colLyncStencils.Masters.Item("Reverse Proxy")
$stnSBASBS = $colLyncStencils.Masters.Item("Branch Office Appliance")
$stnServer = $colLyncStencils.Masters.Item("Server, generic")
$stnStandardFE = $colLyncStencils.Masters.Item("Standard Edition")
$stnVoIPGateway = $colLyncStencils.Masters.Item("Advanced VoIP Gateway")
$stnRoundRectangle = $colBasicStencils.Masters.Item("Rounded Rectangle")

#Create basic background layout in Visio.
#$shpInternetBG = $objPage.Drop($stnRoundRectangle, 1, 2.85)
#$shpInternetBG.Cells("Height").Formula = "5.2 in."
#$shpInternetBG.Cells("Width").Formula = "1.5 in."
#$shpInternetBG.Cells("FillForeGnd").Formula = "RGB(175,0,0)"
#$shpInternetBG.Cells("FillForeGndTrans").Formula = "80%"
#$shpInternetBG.Cells("VerticalAlign").Formula = 0
#$shpInternetBG.Cells("Char.Size").Formula = "14 pt"
#$shpInternetBG.Text = "Internet"

#$shpDMZBG = $objPage.Drop($stnRoundRectangle, 2.5, 2.85)
#$shpDMZBG.Cells("Height").Formula = "5.2 in."
#$shpDMZBG.Cells("Width").Formula = "1.5 in."
#$shpDMZBG.Cells("FillForeGnd").Formula = "RGB(175,175,0)"
#$shpDMZBG.Cells("FillForeGndTrans").Formula = "80%"
#$shpDMZBG.Cells("VerticalAlign").Formula = 0
#$shpDMZBG.Cells("Char.Size").Formula = "14 pt"
#$shpDMZBG.Text = "DMZ"

$LegendText = @()
$ClusterServers
$LegendText += $DefaultSIPDomain + "`nLync Topology`n`n"
$LegendText += $Sites.Count.ToString() + " Sites`n"
$LegendText += $Servers.Count.ToString() + " Servers`n"
$LegendText += $SIPDomains.Count.ToString() + " SIP Domains`n"
if ($Clusters){$LegendText += $Clusters.Count.ToString() + " Clusters`n"}
$LegendText += "`n"
$LegendText += "SIP Domains:`n"
$SIPDomains | foreach-object{$LegendText += $_ + "`n"}
$LegendText += "`n"
if ($Clusters){
$LegendText += "Clusters:`n"
foreach ($Cluster in $Clusters){
$LegendText += $Cluster.Name + "`n"
$Servers | Where-Object{$_.ClusterFQDN -eq $Cluster.Name} | Select $_.Name | foreach-object{
$ClusterServerName = $_.Name
$ClusterServers += $ClusterServerName.SubString(0,$ClusterServerName.IndexOf(".")) + "`n"}
$LegendText += $ClusterServers
$LegendText += "`n"
$ClusterServers = $Null
}}

$shpLegendBG = $objPage.Drop($stnRoundRectangle, 1.75, 4.25)
$shpLegendBG.Cells("Height").Formula = "8 in."
$shpLegendBG.Cells("Width").Formula = "3 in."
$shpLegendBG.Cells("FillForeGnd").Formula = "RGB(230,230,230)"
$shpLegendBG.Cells("FillForeGndTrans").Formula = "80%"
$shpLegendBG.Cells("VerticalAlign").Formula = 0
$shpLegendBG.Cells("Char.Size").Formula = "14 pt"
$shpLegendBG.Text = $LegendText

$shpSiteHeight = (8/$Sites.Count)
$Site
$CurrentSiteNumber = 0
$CurrentServerNumber = 0

foreach ($Site in $Sites){
Write-Host "Building site: "$Site.Name
$CurrentSiteNumber++
$ShapeVerticalCenter = ((($CurrentSiteNumber * (8 / $Sites.Count)) - ((8 / $Sites.Count) / 2))+0.25)
$shpSiteBG = $objPage.Drop($stnRoundRectangle, 7, $ShapeVerticalCenter)
$shpSiteBG.Cells("Height").Formula = $shpSiteHeight.ToString() + " in."
$shpSiteBG.Cells("Width").Formula = "7.5 in."
$shpSiteBG.Cells("FillForeGnd").Formula = "RGB(0,0,150)"
if ($Site.ParentSite -ne $Null) {$shpSiteBG.Cells("FillForeGnd").Formula = "RGB(0,150,150)"}
$shpSiteBG.Cells("FillForeGndTrans").Formula = "80%"
$shpSiteBG.Cells("TextDirection").Formula = 1
$shpSiteBG.Cells("VerticalAlign").Formula = 2
$shpSiteBG.Cells("Char.Size").Formula = "14 pt"
$shpSiteBG.Text = $Site.Name + " : " + $Site.Description

$SiteServers = $Servers | Where-Object {$_.OriginalSiteId -eq $Site.OriginalSiteId}
$SiteClusters = $Clusters | Where-Object {$_.OriginalSiteId -eq $Site.OriginalSiteId}

# foreach ($Server in $Servers){if ($Server.OriginalSiteId -eq $Site.OriginalSiteId){$SiteServerCount++}}
# foreach ($Cluster in $Clusters){if ($Cluster.OriginalSiteId -eq $Site.OriginalSiteId){$SiteClusterCount++}}

foreach ($Server in $SiteServers){
$ServerServices = @()
#$ServerRole = @()
foreach ($Service in $Services){if ($Service.InstalledOn -eq $Server.UniqueID){$ServerServices += $Service.RoleName}}

	$shpServerType = $strServer
	if ($ServerServices -contains "UserServices"){$ServerRole = "Front-End" ; $shpServerType = $stnEnterpriseFE}
	if ($ServerServices -contains "FileStore" -and $ServerServices -notcontains "UserServices"){$ServerRole = "File-Store" ; $shpServerType = $stnFileStore}
	if ($ServerServices -contains "ArchivingServer"){$ServerRole = "Archiving" ; $shpServerType = $stnArchivingServer}
	if ($ServerServices -contains "MonitoringServer"){$ServerRole = "Monitoring" ; $shpServerType = $stnMonitoringServer}
	if ($ServerServices -contains "ConfServices" -and $ServerServices -notcontains "UserServices"){$ServerRole = "AV-Conferencing" ; $shpServerType = $stnAVServer}
	if ($ServerServices -contains "Registrar" -and $ServerServices -notcontains "UserServices"){$ServerRole = "Director" ; $shpServerType = $stnDirector}
	if ($ServerServices -contains "MediationServer" -and $ServerServices -notcontains "UserServices"){$ServerRole = "Mediation" ; $shpServerType = $stnMediationServer}
	if ($ServerServices -contains "Registrar" -and $Site.ParentSite -ne $Null){$ServerRole = "SBA/SBS" ; $shpServerType = $stnSBASBS}
	if ($ServerServices -contains "ExternalServer"){$ServerRole = "External-App Server" ; $shpServerType = $stnApplicationServer}
	if ($ServerServices -contains "Registrar" -and $SiteInfo.Name -like "BackCompatSite"){$ServerRole = "OCS-Server" ; $shpServerType = $stnServer}
	if ($ServerServices -contains "EdgeServer"){$ServerRole = "Edge" ; $shpServerType = $stnEdgeServer}
	if ($ServerServices -contains "PSTNGateway"){$ServerRole = "PSTN-Gateway" ; $shpServerType = $stnVoIPGateway}
	if ($Server.SQLInstances){$Server.SQLInstances.GetEnumerator() | Foreach-Object {if ($ServerServices.Count -eq 0 -and $_.Value.UniqueID -like "sql*"){$ServerRole = "Back-End-SQL" ; $shpServerType = $stnDatabaseServer}}}

	$CurrentServerNumber++
	$ShapeHorizontalSpacing = (7 / (($SiteServers.Count * 0.5) + 1.25))
	$ShapeHorizontalCenter = ($CurrentServerNumber * $ShapeHorizontalSpacing) + 3.25
	
	$RowHeight = 0.2
	if ($CurrentServerNumber -gt [math]::round(($SiteServers.Count / 2),0)){
		$RowHeight = 0.7
		$ShapeHorizontalCenter = (($CurrentServerNumber - [math]::round(($SiteServers.Count / 2),0)) * $ShapeHorizontalSpacing) + 3.25}
	$ShapeVerticalCenter = ((($CurrentSiteNumber * (8 / $Sites.Count))-((8 / $Sites.Count) * $RowHeight)) + 0.25)
	
	
	$shpServer = $objPage.Drop($shpServerType, $ShapeHorizontalCenter, $ShapeVerticalCenter)
	$shpServer.Name = $Server
	$shpLabel = $Server.Name.Split(".")
	$shpServer.Text = $shpLabel[0] + "`n" + $ServerRole
	$shpServer.CellsU("TopMargin").FormulaU = "4 pt"
	$shpServer.CellsU("TxtWidth").FormulaU = "Width * 1.75"
	$shpServer.CellsU("TxtPinY").FormulaU = "Height * -0.5"
	if ($ServerRole -eq "PSTN-Gateway"){$shpServer.CellsU("TxtPinY").FormulaU = "Height * -0.75"}
	$shpServer.CellsU("TxtLocPinY").FormulaU = ""
	$shpServer.CellsU("VerticalAlign").FormulaU = 2
	$shpServer.Cells("Char.Size").Formula = "10 pt"
	$ServerServices = $Null
	$ServerRole = $Null
}
$SiteServerCount = 0
$SiteClusterCount = 0
$CurrentServerNumber = 0

Write-Host "Done!"

}

$objDocument.SaveAs($CurrentDirectory + $VisioFileName) | Out-Null
$colLyncStencils.Close()
$colBasicStencils.Close()
Write-Host "Visio Document Complete!"

}





#Collect TBXML Filename from Command Line Arguments
$TBXMLFileName = $Source
Write-Host "Importing TBXML Config..."
$TBXML = [xml](get-content $TBXMLFileName)
Write-Host "Done"

#Get current working directory and convert it to a string.
$CurrentDirectory = Get-Location
$CurrentDirectory = $CurrentDirectory.ToString()
$CDLastChar = [string]$CurrentDirectory[-1]
if ($CDLastChar -ne "\"){$CurrentDirectory = $CurrentDirectory + "\"}

#Get accepted SIP domains.
$InternalDomainsXML = $TBXML.TopologyBuilder.NewTopology.PartialTopology.InternalDomains
$DefaultSIPDomain = $InternalDomainsXML.DefaultDomain
$SIPDomains = @()
foreach ($SIPDomain in $InternalDomainsXML.InternalDomain) {
$SIPDomains += $SIPDomain.Name
}

#Get URL Configuration.
$URLConfigXML = $TBXML.TopologyBuilder.NewTopology.PartialTopology.SimpleUrlConfiguration.SimpleURL
$URLS = @()
foreach ($URL in $URLConfigXML) {
$objURL =  new-object PSObject -Property @{
Component = ($URL.Component)
ActiveUrl = ($URL.ActiveUrl)
}
$URLS += $objURL
}

#Get Site and Server configurations.
$SitesXML = $TBXML.TopologyBuilder.NewTopology.PartialTopology.CentralSites.CentralSite
$Sites = @()
$CentralSites = @()
$CentralSiteServers = @()
$RemoteSites = @()
$RemoteSiteServers = @()
$Servers = @()
$Clusters = @()
$NetInterfaces = @()
$SQLInstances = @()
foreach ($Location in $SitesXML){
$objSite = new-object PSObject
$objSite | Add-Member -type NoteProperty -name Name -Value $Location.Name."#text"
$objSite | Add-Member -type NoteProperty -name Description -Value $Location.Description."#text"
$objSite | Add-Member -type NoteProperty -name City -Value $Location.Location.City
$objSite | Add-Member -type NoteProperty -name State -Value $Location.Location.State
$objSite | Add-Member -type NoteProperty -name Country -Value $Location.Location.CountryCode
$objSite | Add-Member -type NoteProperty -name OriginalSiteId -Value $Location.OriginalSiteId
$Sites += $objSite
	foreach ($Cluster in $Location.Clusters.Cluster){
		if ($Cluster.IsSingleMachineOnly -eq "false"){
		$objCluster = new-object PSObject -Property @{
		Name = ($Cluster.fqdn);
		UniqueID = ($Cluster.UniqueID);
		OriginalSiteId = ($Cluster.OriginalSiteID);
		ServerNumber = ($Cluster.OriginalNumber);
		IsSingleMachineOnly = ($Cluster.IsSingleMachineOnly);
		SiteName = ($Location.Name."#text");
		}
		$Clusters += $objCluster
		foreach ($Machine in $Cluster.Machines.Machine){
		$objServer = new-object PSObject -Property @{
		Name = ($Machine.fqdn);
		ClusterFQDN = ($Cluster.fqdn);
		UniqueID = ($Cluster.UniqueID);
		OriginalSiteId = ($Cluster.OriginalSiteID);
		OrdinalInCluster = ($Machine.OriginalOrdinalInCluster);
		ServerNumber = ($Cluster.OriginalNumber);
		IsSingleMachineOnly = ($Cluster.IsSingleMachineOnly);
		SiteName = ($Location.Name."#text");
		NetInterfaces = @{}
		}
			$i = 1
			foreach ($NetInterface in $Machine.NetInterface){
				$InterfaceNumber = $NetInterface.InterfaceNumber
				$InterfaceSide = $NetInterface.InterfaceSide
				$IPAddress = $NetInterface.IPAddress
				$ConfiguredIPAddress = $NetInterface.ConfiguredIPAddress
				
				if (!$NetInterface.InterfaceNumber) {$InterfaceNumber = "None"}
				if (!$NetInterface.InterfaceSide) {$InterfaceSide = "None"}
				if (!$NetInterface.IPAddress) {$IPAddress = "None"}
				if (!$NetInterface.ConfiguredIPAddress) {$ConfiguredIPAddress = "None"}			
				
				$objNetInterface = new-object PSObject -Property @{
				ClusterUniqueID = $Machine.UniqueID
				InterfaceNumber = $InterfaceNumber
				InterfaceSide = $InterfaceSide
				IPAddress = $IPAddress
				ConfiguredIPAddress = $ConfiguredIPAddress
				}
				$objServer.NetInterfaces["Interface$i"] = $objNetInterface
				$i++
			}
		$Servers += $objServer
		}}
		else {
		$objServer = new-object PSObject -Property @{
		Name = ($Cluster.fqdn);
		UniqueID = ($Cluster.UniqueID);
		OriginalSiteId = ($Cluster.OriginalSiteID);
		ServerNumber = ($Cluster.OriginalNumber);
		IsSingleMachineOnly = ($Cluster.IsSingleMachineOnly);
		SiteName = ($Location.Name."#text");
		NetInterfaces = @{}
		SQLInstances = @{}
		}
			$i = 1
			foreach ($NetInterface in $Cluster.Machines.Machine.NetInterface){
				$InterfaceNumber = $NetInterface.InterfaceNumber
				$InterfaceSide = $NetInterface.InterfaceSide
				$IPAddress = $NetInterface.IPAddress
				$ConfiguredIPAddress = $NetInterface.ConfiguredIPAddress
				
				if (!$NetInterface.InterfaceNumber) {$InterfaceNumber = "None"}
				if (!$NetInterface.InterfaceSide) {$InterfaceSide = "None"}
				if (!$NetInterface.IPAddress) {$IPAddress = "None"}
				if (!$NetInterface.ConfiguredIPAddress) {$ConfiguredIPAddress = "None"}			
				
				$objNetInterface = new-object PSObject -Property @{
				ClusterUniqueID = $Cluster.UniqueID
				InterfaceNumber = $InterfaceNumber
				InterfaceSide = $InterfaceSide
				IPAddress = $IPAddress
				ConfiguredIPAddress = $ConfiguredIPAddress
				}
				$objServer.NetInterfaces["Interface$i"] = $objNetInterface
				$i++
			}
			$i = 1
			$objSQLInstances = new-object PSObject
			foreach ($SQLInstance in $Cluster.SqlInstances.SqlInstance){
				$SQLName = $SQLInstance.Name
				$SQLUniqueID = $SQLInstance.UniqueID
				$SQLOriginalClusterUniqueID = $SQLInstance.OriginalClusterUniqueID
				
				if (!$SQLInstance.Name){$SQLName = "None"}
				if (!$SQLInstance.UniqueID){$SQLUniqueID = "None"}
				if (!$SQLInstance.OriginalClusterUniqueID){$SQLOriginalClusterUniqueID = "None"}
				
				$objSQLInstance = new-object PSObject -Property @{
				Name = $SQLName
				UniqueID = $SQLUniqueID
				OriginalClusterUniqueID = $SQLOriginalClusterUniqueID
				}
				$objServer.SQLInstances["Instance$i"] = $objSQLInstance
				$i++
			}
	$Servers += $objServer
	}}
	foreach ($RemoteLocation in $Location.RemoteSites.RemoteSite){
	if ($RemoteLocation.Name."#text" -ne $null){
	$objSite = new-object PSObject
	$objSite | Add-Member -type NoteProperty -name Name -Value $RemoteLocation.Name."#text"
	$objSite | Add-Member -type NoteProperty -name Description -Value $RemoteLocation.Description."#text"
	$objSite | Add-Member -type NoteProperty -name City -Value $RemoteLocation.Location.City
	$objSite | Add-Member -type NoteProperty -name State -Value $RemoteLocation.Location.State
	$objSite | Add-Member -type NoteProperty -name Country -Value $RemoteLocation.Location.CountryCode
	$objSite | Add-Member -type NoteProperty -name ParentSite -Value $Location.Name."#text"
	$objSite | Add-Member -type NoteProperty -name OriginalSiteId -Value $RemoteLocation.OriginalSiteId
	$Sites += $objSite
		foreach ($RemoteCluster in $RemoteLocation.Clusters.Cluster){
		$objServer = new-object PSObject -Property @{
		Name = ($RemoteCluster.fqdn);
		UniqueID = ($RemoteCluster.UniqueID);
		OriginalSiteId = ($RemoteCluster.OriginalSiteID);
		ServerNumber = ($RemoteCluster.OriginalNumber);
		IsSingleMachineOnly = ($RemoteCluster.IsSingleMachineOnly);
		SiteName = ($RemoteLocation.Name."#text");
		NetInterfaces = @{}
		SQLInstances = @{}
		}
			$i = 1
			foreach ($NetInterface in $RemoteCluster.Machines.Machine.NetInterface){
				$InterfaceNumber = $NetInterface.InterfaceNumber
				$InterfaceSide = $NetInterface.InterfaceSide
				$IPAddress = $NetInterface.IPAddress
				$ConfiguredIPAddress = $NetInterface.ConfiguredIPAddress
				
				if (!$NetInterface.InterfaceNumber) {$InterfaceNumber = "None"}
				if (!$NetInterface.InterfaceSide) {$InterfaceSide = "None"}
				if (!$NetInterface.IPAddress) {$IPAddress = "None"}
				if (!$NetInterface.ConfiguredIPAddress) {$ConfiguredIPAddress = "None"}			
				
				$objNetInterface = new-object PSObject -Property @{
				ClusterUniqueID = $RemoteCluster.UniqueID
				InterfaceNumber = $InterfaceNumber
				InterfaceSide = $InterfaceSide
				IPAddress = $IPAddress
				ConfiguredIPAddress = $ConfiguredIPAddress
				}
				$objServer.NetInterfaces["Interface$i"] = $objNetInterface
				$i++
			}
			$i = 1
			$objSQLInstances = new-object PSObject
			foreach ($SQLInstance in $RemoteCluster.SqlInstances.SqlInstance){
				$SQLName = $SQLInstance.Name
				$SQLUniqueID = $SQLInstance.UniqueID
				$SQLOriginalClusterUniqueID = $SQLInstance.OriginalClusterUniqueID
				
				if (!$SQLInstance.Name){$SQLName = "None"}
				if (!$SQLInstance.UniqueID){$SQLUniqueID = "None"}
				if (!$SQLInstance.OriginalClusterUniqueID){$SQLOriginalClusterUniqueID = "None"}
				
				$objSQLInstance = new-object PSObject -Property @{
				Name = $SQLName
				UniqueID = $SQLUniqueID
				OriginalClusterUniqueID = $SQLOriginalClusterUniqueID
				}
				$objServer.SQLInstances["Instance$i"] = $objSQLInstance
				$i++
			}
		$Servers += $objServer
		}}
	}}

#Get Service configuration.
$ServicesXML = $TBXML.TopologyBuilder.NewTopology.PartialTopology.Services.Service
$Services = @()
$Dependencies = @()
$Ports = @()
foreach ($Service in $ServicesXML){
	$objService = new-object PSObject -Property @{
		Type = ($Service.Type)
		UniqueId = ($Service.UniqueId)
		InstalledOn = ($Service.InstalledOn)
		RoleName = ($Service.RoleName)
		OriginalSiteId = ($Service.OriginalSiteId)
		OriginalInstance = ($Service.OriginalInstance)
		Dependencies = @()
		Ports = @{}
	}

foreach ($Dependency in $Service.DependsOn.Dependency){
	if (!$Dependency.ServiceUniqueId){$objService.Dependencies += "None"}
	else {$objService.Dependencies += $Dependency.ServiceUniqueId}}

foreach ($Port in $Service.Ports.Port){
	if (!$Port.Owner){$Owner = " "} else {$Owner = $Port.Owner}
	if (!$Port.Usage){$Usage = " "} else {$Usage = $Port.Usage}
	if (!$Port.InterfaceSide){$InterfaceSide = " "} else {$InterfaceSide = $Port.InterfaceSide}
	if (!$Port.InterfaceNumber){$InterfaceNumber = " "} else {$InterfaceNumber = $Port.InterfaceNumber}
	if (!$Port.Port){$PortNumber = " "} else {$PortNumber = $Port.Port}
	if (!$Port.Protocol){$Protocol = " "} else {$Protocol = $Port.Protocol}
	if (!$Port.UrlPath){$UrlPath = " "} else {$UrlPath = $Port.UrlPath}
	if (!$Port.AuthorizesRequests){$AuthorizesRequests = " "} else {$AuthorizesRequests = $Port.AuthorizesRequests}
	if (!$Port.Range){$Range = " "} else {$Range = $Port.Range}
	if (!$Port.GruuType){$GruuType = " "} else {$GruuType = $Port.GruuType}
	if (!$Port.ConfiguredFqdn){$ConfiguredFqdn = " "} else {$ConfiguredFqdn = $Port.ConfiguredFqdn} 
	
	$objPort = new-object PSObject -Property @{
		Owner = $Owner
		Usage = $Usage
		InterfaceSide = $InterfaceSide
		InterfaceNumber = $InterfaceNumber
		PortNumber = $PortNumber
		Protocol = $Protocol
		UrlPath = $UrlPath
		AuthorizesRequests = $AuthorizesRequests
		Range = $Range
		GruuType = $GruuType
		ConfiguredFqdn = $ConfiguredFqdn
	}
$objService.Ports["Port $PortNumber"] = $objPort
}
$Services += $objService
}

# Determine which document type to produce from command line switch.
if ($xls){
	Build-Spreadsheet
}
if ($vsd){
	Build-Drawing
}