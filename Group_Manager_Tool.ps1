Function Welcome{

         Clear;       
         Write-Host "***********************************************"  -foreground Red 
         Write-Host "     Exchange Security Group Tool      " -foreground Red 
         Write-Host "***********************************************"  -foreground Red 
}

Function Login {

        while(1){
                Welcome;
                Import-Module MsOnline;
                Write-Host "step 1" -ForegroundColor yellow;
                Write-Host " Enter Office365 account : " -nonewline
                $global:adm_account = Read-Host ;
                Write-Host "--------------------------------------------------"-ForegroundColor yellow;

                Write-Host "step 2" -ForegroundColor yellow;
                Write-Host " Enter Password : " -nonewline
                $global:adm_password_plain = Read-Host 
                $global:adm_password_encrypt = convertto-securestring  $adm_password_plain -asplaintext -force

                $global:adm_cred=New-Object System.Management.Automation.PSCredential($adm_account,$adm_password_encrypt);
                $report = Connect-MsolService -credential $adm_cred 2>&1;
                $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
				                if($err){
					                Write-Host " $report" -background Black -foreground Red	
					                Read-Host;
                                    Clear;
				                }
				                else{
					                Write-Host "Login Success!" -background Black -foreground Magenta	
                                    Write-Host "--------------------------------------------------"-ForegroundColor yellow;
					                Break;
				                }
        }    
} 
Function Import_Exchange_Module {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $adm_cred -Authentication Basic -AllowRedirection
        Import-PSSession $Session
        Start-Sleep -s 1
}

Function Print_Choose{
        Write-Host " 1.Get Security Group" -ForegroundColor yellow;
        Write-Host " 2.Create Security Group" -ForegroundColor yellow;
        Write-Host " 3.Remove Security Group" -ForegroundColor yellow;
        Write-Host " 4.Get Group Member" -ForegroundColor yellow;
        Write-Host " 5.Add Group Member" -ForegroundColor yellow; 
        Write-Host " 6.Remove Group Member" -ForegroundColor yellow;
        
        Write-Host " 9.Exit" -ForegroundColor yellow;
}

Function Get_Security_Group {
        Get-DistributionGroup | Out-GridView -Title "Security Group " ;
}

Function Create_Security_Group {
        Clear;
        Welcome;
        Write-Host "Login : $adm_account" -background Black -foreground Magenta
        Write-Host "Create Security Group" -ForegroundColor yellow;
            $Group_name = Read-Host " Enter Group Name "
            $Group_mail = Read-Host " Enter Group Mail ( Ex: class1@nctu.edu.tw ) "
        New-DistributionGroup -Name $Group_name -PrimarySmtpAddress $Group_mail -Type "Security" -IgnoreNamingPolicy ;
}

Function Remove_Security_Group {
        Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to delete"  -PassThru | ForEach-Object { 
                Remove-DistributionGroup -Identity $_.Displayname -Confirm:$false ; 
                Write-Host "Remove Group *$_.Displayname* Success" -background Black -foreground Green;
        };
}

Function Get_Group_Member {
        Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to get member" -PassThru | ForEach-Object {
            $GroupName =  $_.Displayname;
            Get-DistributionGroupMember -Identity $GroupName | Out-Gridview -Title "Group *$GroupName* Member"; 
        };
}

Function Add_Group_Member {
        Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to add member" -PassThru | ForEach-Object {
            $GroupName =  $_.Displayname;
            Get-MsolUser -All  | Out-Gridview -Title "Choose the Member you want to add to *$GroupName*" -PassThru | ForEach-Object {
                    Add-DistributionGroupMember -Identity $GroupName -Member $_.UserPrincipalName -BypassSecurityGroupManagerCheck;
            }
            Get-DistributionGroupMember -Identity $GroupName | Out-Gridview -Title "Group *$GroupName* Member"; 
        };
}

Function Remove_Group_Member {
        Get-DistributionGroup | Out-Gridview -Title "Choose the Group you want to delete member" -PassThru | ForEach-Object { 
                $Group_Name = $_.Displayname; 
                $Result = Get-DistributionGroupMember -Identity $Group_Name | Out-Gridview -Title "Choose Group *$GroupName* Member to delete" -PassThru ;
                 Foreach( $Each in $Result){
                           Remove-DistributionGroupMember -Identity $Group_Name  -Member $Each.Name -BypassSecurityGroupManagerCheck -Confirm:$false;
                           Write-Host "Remove Member *$Each* from *$Group_Name* Success" -background Black -foreground Green;  
                 };
        };



}


<# Main #>
    
Login;
Import_Exchange_Module;

while(1){
            Welcome;
            Write-Host "Login : $adm_account" -background Black -foreground Magenta
            Print_Choose;
            $choose = Read-Host "Please choose the number "
			Switch ($choose) { 
					1{ Get_Security_Group; } 
					2{ Create_Security_Group; }
                    3{ Remove_Security_Group; }
                    4{ Get_Group_Member;}
                    5{ Add_Group_Member;}
                    6{ Remove_Group_Member;}
                    9{ }
					default { ; }
			}
			if($choose -ne 9){
                    Write-Host "Press any key to continue" -ForegroundColor Red;
					Read-Host;
			}
            else{
                    break;
            }       
}

<# End Main #>