#v1.0 vMan.ch, 06.04.2021 - Quick And Dirty Version

<#
.SYNOPSIS
Uses RVTools vInfo extract data from all vCenter's specified in VCs and Pushes data directly to PowerBI Dataset Tables with TimeStamp when the Script was run.
	
    Requires following module

    Install-Module -Name import-excel -Scope AllUsers

    Requires RVTools

    https://robware.net/
#>

param
(
    [Array]$VCs = 'vcsa.vman.ch',
    [String]$RVUser = 'rv_tools@vman.sso',
    [String]$RVEncryptedPassword = '_RVToolsPWDZZZZZZZZZZYYYYYYYYYYYYYYYYYYYYYYYPPPPPPPPPOOOO=',
    [String]$RVExtract = "ExportvInfo2xlsx",
    [String]$RvToolsExtract = "RVToolsPowerBIUpload.xlsx",
    [String]$PowerBIDatasetName = "RVTools",
    [String]$PowerBIDatasetTable = "vInfo",
    [String]$PowerBIGroupID = "1y123x45-z67b-0000-0000-zzz111z222222",
    [String]$PowerBIClientID = "bbbccc111-2z2z-111z-aaa6-9bbb1bbb11b1",
    [String]$PowerBICertThumb = "1111aaaa111111a1111111a11a1a1a111a111aa1a",
    [String]$PowerBITenant = "2a222222-a2a2-2a22-aaa2-11aaaaaaaa11",
    [Bool]$CreateDataset = $false
)

#Variables
$MasterTable = @()
$RunDateTime = (Get-date)
[String] $RVToolsPath = "C:\Program Files (x86)\Robware\RVTools"
[String] $RVOutputFolder = "C:\temp"
# cd to RVTools directory
set-location $RVToolsPath

#Connect to PowerBI
Write-Host "Open Connection to Connect to PowerBI" -ForegroundColor Green

$PowerBIConnection = Connect-PowerBIServiceAccount -ServicePrincipal -ApplicationId $PowerBIClientID -CertificateThumbprint $PowerBICertThumb -Tenant $PowerBITenant

If ($PowerBIConnection) {

    ForEach ($VC in $VCs){

        # Run RVTools
        Write-Host "Start export for vCenter $VC" -ForegroundColor DarkYellow
        $Arguments = "-u $RVUser -p $RVEncryptedPassword -s $VC -c ExportAll2xlsx -d $RVOutputFolder -f $RvToolsExtract -DBColumnNames -ExcludeCustomAnnotations"
        $Process = Start-Process -FilePath ".\RVTools.exe" -ArgumentList $Arguments -NoNewWindow -Wait -PassThru

        If ($Process.ExitCode -eq -1){
                Write-Host "Error: Export failed! RVTools returned exitcode -1, probably a connection error! Script stopped" -ForegroundColor Red
                exit 1
            }

        switch ($PowerBIDatasetTable)
        {
            vInfo {

                $MasterTable = import-excel -path "$RVOutputFolder\$RvToolsExtract" -WorksheetName vInfo

                $MasterTable | ForEach { Add-Member -InputObject $_ -Name 'Timestamp' -MemberType noteproperty –Value $RunDateTime }

                If ($CreateDataset){

                    Write-Host "Creating Dataset and vInfo table in PowerBI for the 1st time" -ForegroundColor Green

                        $col0 = New-PowerBIcolumn -Name "TimeStamp" -DataType DateTime
                        $col1 = New-PowerBIColumn -Name "vInfoVMName" -DataType String
                        $col2 = New-PowerBIColumn -Name "vInfoPowerstate" -DataType String
                        $col3 = New-PowerBIColumn -Name "vInfoTemplate" -DataType String
                        $col4 = New-PowerBIColumn -Name "vInfoGuestHostName" -DataType String
                        $col5 = New-PowerBIColumn -Name "vInfoGueststate" -DataType String
                        $col6 = New-PowerBIColumn -Name "vInfoCPUs" -DataType String
                        $col7 = New-PowerBIColumn -Name "vInfoMemory" -DataType String
                        $col8 = New-PowerBIColumn -Name "vInfoPrimaryIPAddress" -DataType String
                        $col9 = New-PowerBIColumn -Name "vInfoInUse" -DataType String
                        $col10 = New-PowerBIColumn -Name "vInfoProvisioned" -DataType String
                        $col11 = New-PowerBIColumn -Name "vInfoOSTools" -DataType String
                        $col12 = New-PowerBIColumn -Name "vInfoHost" -DataType String
                        $col13 = New-PowerBIColumn -Name "vInfoVISDKServer" -DataType String
                        $col14 = New-PowerBIColumn -Name "vInfoCluster" -DataType String

                    $vInfoTable = New-PowerBITable -Name $PowerBIDatasetTable -Columns $col0,$col1,$col2,$col3,$col4,$col5,$col6,$col7,$col8,$col9,$col10,$col11,$col12,$col13,$col14

                    # Create dataset
                    $RVdataset = New-PowerBIDataSet -Name $PowerBIDatasetName -Tables $vInfoTable

                    # Add to workspace
                    $RVdatasetid = Add-PowerBIDataSet -DataSet $RVdataset -WorkspaceId $PowerBIGroupID

                    $PowerBIDatasetID = $RVdatasetid.id

                    Clear-Variable vInfoTable,RVdataset,RVdatasetid

                } else {

                    $MasterLookupTable = $MasterTable | select TimeStamp,vInfoVMName,vInfoPowerstate,vInfoTemplate,vInfoGuestHostName,vInfoGueststate,vInfoCPUs,vInfoMemory,vInfoPrimaryIPAddress,vInfoInUse,vInfoProvisioned,vInfoOSTools,vInfoHost,vInfoVISDKServer,vInfoCluster


                            If ($PowerBIDatasetID){

                                Add-PowerBIRow -DataSetId $PowerBIDatasetID -WorkspaceId $PowerBIGroupID -TableName $PowerBIDatasetTable -Rows $MasterLookupTable

                            }else{

                                $PowerBIDatasetID = Get-PowerBIDataset -WorkspaceId $PowerBIGroupID | Where name -eq $PowerBIDatasetName | Select Id

                                Add-PowerBIRow -DataSetId $PowerBIDatasetID.Id -WorkspaceId $PowerBIGroupID -TableName $PowerBIDatasetTable -Rows $MasterLookupTable

                            }
        
                        }

               Remove-Item -Path "$RVOutputFolder\$RvToolsExtract" -Force

               Clear-Variable Arguments,Process,VC,MasterTable,MasterLookupTable

            }

            vHost {}

            vCluster {}

        }
    }
}else{

    Write-Host "Failed to connect to PowerBI, quitting script" -ForegroundColor Red
    Exit 1

}

Disconnect-PowerBIServiceAccount
