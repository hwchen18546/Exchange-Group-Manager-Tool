Function Welcome{

         Clear;       
         Write-Host "*******************************************************"  -foreground Red;
         Write-Host "      Exchange Online Security Group Manager Tool      "  -foreground Red;
         Write-Host "*******************************************************"  -foreground Red;
}

Function Login {

        while(1){
                Welcome;
                Import-Module MsOnline;
                Write-Host "step 1" -ForegroundColor yellow;
                Write-Host " Enter Office365 account : " -nonewline;
                $global:adm_account = Read-Host ;
                Write-Host "--------------------------------------------------"-ForegroundColor yellow;

                Write-Host "step 2" -ForegroundColor yellow;
                Write-Host " Please enter your password : " -nonewline;
				$global:adm_password_plain = Read-Host;
                $global:adm_password_encrypt = convertto-securestring  $adm_password_plain -asplaintext -force;

                $global:adm_cred=New-Object System.Management.Automation.PSCredential($adm_account,$adm_password_encrypt);
                $report = Connect-MsolService -credential $adm_cred 2>&1;
                $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
				                if($err){
					                Write-Host " $report" -background Black -foreground Red	;
					                Read-Host;
                                    Clear;
				                }
				                else{
					                Write-Host "Login Success!" -background Black -foreground Magenta;	
                                    Write-Host "--------------------------------------------------"-ForegroundColor yellow;
					                Break;
				                }
        }    
} 

Function Import_Exchange_Module {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $adm_cred -Authentication Basic -AllowRedirection;
        Import-PSSession $Session;
}

Function Print_Choose{
        Write-Host " 1.Get Security Group" -ForegroundColor yellow;
        Write-Host " 2.Create Security Group" -ForegroundColor yellow;
        Write-Host " 3.Remove Security Group" -ForegroundColor yellow;
        Write-Host " 4.Get Group Member" -ForegroundColor yellow;
        Write-Host " 5.Add Group Member" -ForegroundColor yellow; 
        Write-Host " 6.Remove Group Member" -ForegroundColor yellow;
        Write-Host " 7.Import-Csv Create Group" -ForegroundColor yellow;
        Write-Host " 8.Import-Csv Add Group Member" -ForegroundColor yellow;
        Write-Host " 9.Import-Csv Group Member List" -ForegroundColor yellow;
        Write-Host " 0.Exit" -ForegroundColor yellow;
}

Function Get_Security_Group {
        Get-DistributionGroup | Out-GridView -Title "Choose the Security Group" -PassThru | Format-List;
}

Function Create_Security_Group {
        Clear;
        Welcome;
        Write-Host "Login : $adm_account" -background Black -foreground Magenta;
        Write-Host "Create Security Group" -ForegroundColor yellow;
            $Group_name = Read-Host " Enter Group Name ";
            $split = $adm_account.split("@");
            $Group_address = $Group_name+"@"+$split[1];
        New-DistributionGroup -Name $Group_name -PrimarySmtpAddress $Group_address -Type "Security" -IgnoreNamingPolicy ;
        Set-DistributionGroup -Identity $Group_name -RequireSenderAuthenticationEnabled $false;
}

Function Remove_Security_Group {
        Write-Host "You are ready to delete Group. Press 'y' to continue" -NoNewline -background Black -foreground Magenta;
        $Confirm = Read-Host " ";
        if($Confirm -eq "y" -Or $Confirm -eq "Y"){
            Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to delete"  -PassThru | ForEach-Object { 
                    Remove-DistributionGroup -Identity $_.Displayname -Confirm:$false -BypassSecurityGroupManagerCheck;
                    $Delete_Group = $_.Displayname;
                    Write-Host "Remove Group *$Delete_Group* Success" -background Black -foreground Green;
            };
        }
}

Function Get_Group_Member {
        Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to GET member" -PassThru | ForEach-Object {
            $GroupName =  $_.Displayname;
            Get-DistributionGroupMember -Identity $GroupName | Out-Gridview -Title "Group *$GroupName* Member"; 
        };
}

Function Add_Group_Member {
        Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to ADD member" -PassThru | ForEach-Object {
            $GroupName =  $_.Displayname;
            Get-MsolUser -All  | Out-Gridview -Title "Choose the Member you want to add to *$GroupName*" -PassThru | ForEach-Object {
                    Add-DistributionGroupMember -Identity $GroupName -Member $_.UserPrincipalName -BypassSecurityGroupManagerCheck;
            }
            Get-DistributionGroupMember -Identity $GroupName | Out-Gridview -Title "Group *$GroupName* Member"; 
        };
}

Function Remove_Group_Member {
        Write-Host "You are ready to delete Group member. Press 'y' to continue" -NoNewline -background Black -foreground Magenta;
        $Confirm = Read-Host " ";
        if($Confirm -eq "y" -Or $Confirm -eq "Y"){
            Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to DELETE member" -PassThru | ForEach-Object { 
                    $Group_Name = $_.Displayname;
                    $Result = Get-DistributionGroupMember -Identity $Group_Name | Out-Gridview -Title "Choose Group *$GroupName* Member to delete" -PassThru ;
                     Foreach( $Each in $Result){
                               Remove-DistributionGroupMember -Identity $Group_Name  -Member $Each.Name -BypassSecurityGroupManagerCheck -Confirm:$false;
                               Write-Host "Remove Member *$Each* from *$Group_Name* Success" -background Black -foreground Green;  
                     };
            };
        }


}

Function Import_CSV_Create_Group {
    $split = $adm_account.split("@");
    Get-DistributionGroup | Select-Object -Property SamAccountName -First 1 | Export-Csv sample_group.csv -encoding "utf8"; #sample
    $CurrentPath = $(Get-Location).ToString();
    Write-Host "We create a reference sample at $CurrentPath\sample_group.csv"  -background Black -foreground Magenta;
    Write-Host "Do you want to open it?(Y/N): " -NoNewline;
    $open_CSV = Read-Host;
    if($open_CSV -ne "N" -or $open_CSV -ne "n"){
          Invoke-Item sample_group.csv;
    }
    while(1){
            $Importfile = Read-Host " Please enter the Csv File name ";
            $report =  Get-Content -path $Importfile  2>&1;
            $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"} 
			             if($err){
					              Write-Host " Can't open file." -background Black -foreground Red;	
			             }
			             else{
                                 Write-Host " Import $Importfile ..." -background Black -foreground Green;	
                                 Break;
			             }
    }
    Import-Csv -path $Importfile | ForEach-Object{
        $Group_name = $_.SamAccountName; 
        $Group_address = $Group_name+"@"+$split[1];
        $report = New-DistributionGroup -Name $Group_name -PrimarySmtpAddress $Group_address -Type "Security" -IgnoreNamingPolicy;
        Set-DistributionGroup -Identity $Group_name -RequireSenderAuthenticationEnabled $false;   
    }
    Write-Host "Import Done!" -background Black -foreground Green;
}

Function Import_CSV_Add_Member {

    Get-MsolUser -All | Select-Object -Property UserPrincipalName -Last 1 | Export-Csv sample_member.csv -encoding "utf8"; #sample
    $CurrentPath = $(Get-Location).ToString();
    Write-Host "We create a reference sample at $CurrentPath\sample_member.csv"  -background Black -foreground Magenta;
    Write-Host "Do you want to open it?(Y/N): " -NoNewline;
    $open_CSV = Read-Host;
    if($open_CSV -ne "N" -or $open_CSV -ne "n"){
          Invoke-Item sample_member.csv;
    }
    while(1){
            $Importfile = Read-Host " Please enter the Csv file name ";
            $report =  Get-Content -path $Importfile  2>&1;
            $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"} 
			             if($err){
					              Write-Host " Can't open file." -background Black -foreground Red;	
			             }
			             else{
                                 Write-Host " Import $Importfile ..." -background Black -foreground Green;
                                 Break;
			             }
    }
    Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to ADD member" -PassThru | ForEach-Object {
            $Group_name =  $_.Displayname;
    }; 
    Import-Csv -path $Importfile | ForEach-Object{
            $member_name = $_.UserPrincipalName; 
            Add-DistributionGroupMember -Identity $Group_name -Member $member_name -BypassSecurityGroupManagerCheck;  
    };
    Get-DistributionGroupMember -Identity $Group_name | Out-Gridview -Title "Group *$Group_name* Member";
    Write-Host "Import Done!" -background Black -foreground Green;
}

Function Import_CSV_Group_Member_List {
    Get-MsolUser -All | Select-Object -Property UserPrincipalName,GroupName -Last 1 | Export-Csv sample_list.csv -encoding "utf8"; #sample
    $CurrentPath = $(Get-Location).ToString();
    Write-Host "We create a reference sample at $CurrentPath\sample_list.csv"  -background Black -foreground Magenta;
    Write-Host "Do you want to open it?(Y/N): " -NoNewline;
    $open_CSV = Read-Host;
    if($open_CSV -ne "N" -or $open_CSV -ne "n"){
          Invoke-Item sample_list.csv;
    }
    while(1){
            $Importfile = Read-Host " Please enter the Csv file name ";
            $report =  Get-Content -path $Importfile  2>&1;
            $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"} 
			             if($err){
					              Write-Host " Can't open file." -background Black -foreground Red;	
			             }
			             else{
                                 Write-Host " Import $Importfile ..." -background Black -foreground Green;
                                 Break;
			             }
    }
    Import-Csv -path $Importfile | ForEach-Object{
            $member_name = $_.UserPrincipalName; 
            $Group_name = $_.GroupName;
            Add-DistributionGroupMember -Identity $Group_name -Member $member_name -BypassSecurityGroupManagerCheck;  
    };
    Write-Host "Import Done!" -background Black -foreground Green;
}

<# Main #>
    
Login;
Import_Exchange_Module;

while(1){
            Welcome;
            Write-Host "Login : $adm_account" -background Black -foreground Magenta;
            Print_Choose;
            $choose = Read-Host "Please choose the number "
			Switch ($choose) { 
					1{ Get_Security_Group; } 
					2{ Create_Security_Group; }
                    3{ Remove_Security_Group; }
                    4{ Get_Group_Member;}
                    5{ Add_Group_Member;}
                    6{ Remove_Group_Member;}
                    7{ Import_CSV_Create_Group;}
                    8{ Import_CSV_Add_Member;}
                    9{ Import_CSV_Group_Member_List;}
					default { ; }
			}
			if($choose -ne -0){
                    Write-Host "Press any key to continue" -ForegroundColor Red;
					Read-Host;
			}
            else{
                    Get-PSSession | Remove-PSSession;
                    break;
            }       
}

<# End Main #>