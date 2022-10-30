#Create AD User
#change lower case to upper case e.g. bipin lamichhane TO Bipin Lamichhane
function Naming_Case($first, $second){
$firstN = $first.substring(0,1).ToUpper() + $first.Substring(1).ToLower()
$secondN = $second.substring(0,1).ToUpper() + $second.Substring(1).ToLower()
return @("$firstN", "$secondN")
}

#Check user in AD or not and pass it to create Unique ID for new user
function Check_User($samname){
    try{ 
        $ad_user = Get-ADUser -Identity $samname -Properties SamAccountName -ErrorAction Stop
        $aduser = $ad_user.SamAccountName
        if($samname -eq $aduser){
        return $true
        }
        else{
        return $false
        }
       }catch{
        Write-host "Created Unique ID" -ForegroundColor Green
        return $false
        }     
    }
#Copy Group from one user to another user
function Copy_group($FirstID, $SecondID){
    Write-Host "You are about to copy AD groups" -BackgroundColor Red
    $ask = 'y' #Read-Host "Would you like to copy AD of $SecondID to $FirstID (Y/N) " 
    if($ask -eq "Y" -or $ask -eq "y"){
     try{  
     
     
        #Get-ADUser -Identity $FirstID -Properties MemberOf | ForEach-Object {
        #$_.MemberOf | 
        #Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false -verbose 
       # write-host "REMOVED AD Groups: $A"  -ForegroundColor red 
        #} 
        $referenceuser = Get-ADUser -Identity $SecondID -Properties MemberOf
        $groups= $referenceuser.MemberOf
        $groups | Add-ADGroupMember -Members $FirstID -Verbose
        Write-Host "Groups has been copied from $SecondID to $FirstID " -BackgroundColor DarkGray
    }
    catch{
    Write-Warning "You need higher access" 
    }
    }
    else{
    Write-Host "AD has not been copied"
    }
}
#collecting information from existing user
function Get_Info($userID){
try{
    $sam_name = $userID
    $user_property = Get-ADUser -Identity $sam_name -Properties *
    [String]$email = $user_property.EmailAddress.split("@")[1]
    if($user_property.Manager -eq $null){
    [String]$manager = Read-host " Enter Manager FUll Name"
    $sammy = get-aduser -Filter {Name -eq $manager} -Properties *
    $samName =$sammy.SamAccountName
    Write-Host $samName
    }

    else{
    [String]$manager = $user_property.Manager.split( ",")[0].split( "=")[1]
    
    $sammy = get-aduser -Filter {Name -eq $manager} -Properties *
    $samName =$sammy.SamAccountName 
    }
    $name = $user_property.CN
    [String]$location = $user_property.DistinguishedName -replace ("CN=" + "$name,")
    [String]$displayN = $user_property.DisplayName.split("(")[-1]
    return @("$email", "$SamName", "$location", "$displayN") 
        }catch{
        Write-Warning " Incorrect ID, Try Again " 
        }
}
#User Details and create AD
function User_AD{
        [cmdletBinding()]
        param(

            [Parameter(Mandatory)]
            [String] $firstName,
            [Parameter(Mandatory)]
            [String] $lastName,
            [Parameter(Mandatory)]
            [String] $samName,
            [Parameter(Mandatory)]
            [String] $title,
            [Parameter(Mandatory)]
            [String] $Departname,
            [Parameter(Mandatory)]
            [String] $manager,
            [Parameter(Mandatory)]
            [String] $email,
            [Parameter(Mandatory)]
            [String] $path,
            [Parameter(Mandatory)]
            [String] $UPN
            )
            $full_Name = "$firstName" + " " + "$lastName"
            $display_Name = "$lastName"+", "+"$firstName "+"$UPN"
            #$home = "\\caunsfs01\users\"+"$samName"
            $password = ConvertTo-SecureString -String "P@ssT6word123" -AsPlainText -Force        
            New-ADUser -Name "$full_Name" -GivenName "$firstName" -Surname "$lastName" -Title "$title" -SamAccountName $samName -Department "$Departname" -Manager "$manager" -EmailAddress "$email" -DisplayName "$display_Name" -AccountPassword $password -Enabled $true -Path "$path" -ChangePasswordAtLogon $true -UserPrincipalName "$email" -Verbose
            Set-ADUser $samName -Description $title -Verbose
            Set-ADUser $samName -Office "Chatswood" -Verbose
            if($UPN -eq "(HAL AU)" ){
            Set-ADUser $samName -Company "Holland America Line" -Verbose
            }
            else{
            Set-ADUser $samName -Company "Carnival Australia" -Verbose
            }
                                             
}
#create unique ID
function Create_UniqueID($fname , $lname){
        #Create User ID / SamAccountName
        #Making Unique samName
        Write-Host "----------------------------------------------" -BackgroundColor DarkGray
        $f_Name = $fname
        $l_Name = $lname
        $samname1 = $f_Name + $l_Name[0]
        $samname2 = $f_Name + $l_Name[0] + $l_Name[1]
        $samname3 = $f_Name + $l_Name[0] + $l_Name[1] + $l_Name[2]
        $AD_check = Check_User -samname $samname1
        if($AD_check -eq $true){
            $AD_check2 = Check_User -samname $samname2
            if($AD_check2 -eq $true){
                return $samname3
                Write-Warning "There are 2 userID already on same_name(ignore this warning)"
                }
            else{
            return $samname2
            Write-Warning "There is 1 userID already on same name"
            }
        }
        else{
        return $samname1
        Write-Host "This is unique ID"
        } 
}

#Get New User Details
function user_details{
$fname = Read-Host "Enter User First Name"
$lname = Read-host "Enter User Last Name" 
$title = Read-host "Enter Title"
$depM  = Read-host "Enter department"
$fixed_Names = Naming_case -first $fname -second $lname
$f_name = $fixed_Names[0]
$l_name = $fixed_Names[1]
return @("$f_name", "$l_name", "$title", "$depM")
}


#Default AD group that is for everyone like ExchangeOnline...
function default_groups($samname){
$defaul_groups = @("CAU - M365-LIC (E5 Suite)", "CAU - intune - shoreside users")
foreach($i in $defaul_groups){
Get-ADGroup -Filter 'Name -eq $i' | Add-ADGroupMember -Members $samname -ErrorAction Ignore
}
}
#Remove group from AD
function remove_groups($samname){
$remove_groups = @("")
foreach($i in $remove_groups){
Remove-adgroupmember -identity $i -members $samname -Confirm:$false
}
}
#Excel for new starter into G:\Information Technology\Service Desk\Checklist\New starter Checklist
function create_excel{
[cmdletBinding()]
        param(

            [Parameter(Mandatory)]
            [String] $firstName,
            [Parameter(Mandatory)]
            [String] $lastName,
            [Parameter(Mandatory)]
            [String] $samName
            )
            $fullname = $firstName +" "+$lastName

           try{ 
            $XL = New-Object -ComObject Excel.Application
            $XL.Visible =$true

            $starter = $XL.Workbooks.Open("G:\Information Technology\Service Desk\Checklist\New starter Checklist\Do_Not_Edit.xlsx")
            $starter.ActiveSheet.Cells.Item(3,3) = $fullname
            $starter.ActiveSheet.Cells.Item(4,3) = $samName
            $starter.ActiveSheet.Cells.Item(10,4) = "Completed"
            $starter.ActiveSheet.Cells.Item(11,4) = "Completed"
            $starter.ActiveSheet.Cells.Item(12,4) = "Completed"
            $starter.ActiveSheet.Cells.Item(13,4) = "Completed"
            $starter.ActiveSheet.Cells.Item(18,4) = "Completed"
            $starter.ActiveSheet.Cells.Item(37,4) = "Completed"
            $starter.ActiveSheet.Cells.Item(41,4) = $env:username
            $XL.DisplayAlerts = 'False'
            $ext="_New starter.xlsx"
            $path="G:\Information Technology\Service Desk\Checklist\New starter Checklist\$samName$ext"
            $starter.SaveAs($path) 
            $starter.Close
            $XL.DisplayAlerts = 'False'
            $XL.Quit() 
            Write-Host "Excel $samName$ext has been created"
            }catch{
            Write-Warning "Excel is not created"
            }

}
#Termination excel sheet for user into "G:\Information Technology\Service Desk\Checklist\Termination checklist
function ter_excel{
[cmdletBinding()]
        param(

            [Parameter(Mandatory)]
            [String] $firstName,
            [Parameter(Mandatory)]
            [String] $lastName,
            [Parameter(Mandatory)]
            [String] $depart,
            [Parameter(Mandatory)]
            [String] $Manager,
            [Parameter(Mandatory)]
            [String] $time,
            [Parameter(Mandatory)]
            [String] $SamName
            )
            $fullname = $firstName +" "+$lastName

           try{ 
            $XL = New-Object -ComObject Excel.Application
            $XL.Visible =$true

            $starter = $XL.Workbooks.Open("G:\Information Technology\Service Desk\Checklist\Termination checklist\Do_not_edit.xlsx")
            $starter.ActiveSheet.Cells.Item(7,6) = $fullname
            $starter.ActiveSheet.Cells.Item(8,6) = $manager
            $starter.ActiveSheet.Cells.Item(9,6) = $time
            $starter.ActiveSheet.Cells.Item(11,6) = $depart
            $starter.ActiveSheet.Cells.Item(12,6) = "yes"
            $starter.ActiveSheet.Cells.Item(57,7) = $env:username
            $XL.DisplayAlerts = 'False'
            $ext="_termination.xlsx"
            $path="G:\Information Technology\Service Desk\Checklist\Termination checklist\$samName$ext"
            $starter.SaveAs($path) 
            $starter.Close
            $XL.DisplayAlerts = 'False'
            $XL.Quit() 
            Write-Host "Excel $samName$ext has been created"

            }catch{
            Write-Warning "Excel is not created"
            }

}

#Add additional information into AD
function default_attributes($sam_Name, $email ){
    $email1= $email.split("@")[0]+ ".cau@carnivalcorp.mail.onmicrosoft.com"
    Set-ADUser "$sam_Name" -Add @{proxyAddresses="smtp:$email1"} -Verbose 
    Set-ADUser "$sam_Name" -Add @{proxyAddresses="SMTP:$email"} -Verbose
    Set-ADUser "$sam_Name" -Add @{extensionAttribute5="O365TEAMS"} -Verbose
    #Set-ADUser -Identity $sam_Name -HomeDirectory "\\caunsfs01\users\$sam_Name" -HomeDrive H: -Verbose         
}

#Changing value/name of AD user
#Accessing AD object and changing the Name, email surname, proxyAddress
function change_name{
[cmdletBinding()]
        param(

            [Parameter(Mandatory)]
            [String] $firstName,
            [Parameter(Mandatory)]
            [String] $lastName,
            [Parameter(Mandatory)]
            [String] $samName
            )

            try{
            #Changing object Name
            Get-ADUser $samName | Rename-ADObject -NewName $firstName" "$lastName
            #Changinf first Name and Last Name inside the object
            Set-ADUser $samName -GivenName $firstName -Surname $lastName -Verbose -ErrorAction Ignore

            #getting full email Address
            $sammy = Get-ADUser -Identity $samName -Properties *
            $oldemail = $sammy.EmailAddress

            #getting current email domain and display UPN 
            $info = Get_Info -userID $samName
            [String]$email= $info[0]
            [String]$displayName= $info[3]

            #creating new email address
            $new_email = "$firstName"+"."+"$lastName"+"@"+"$email"
            
            #setting with new email, calling for function that changes everything required after change email
            change_email -oldemail $oldemail -newemail $new_email -samName $samName
            
            #changing display name
            displayName -fname $firstName -lname $lastName -samName $samName -upn1 $displayName

            #create new samName
            $new_sam= Create_UniqueID -fname $FirstName -lname $lastName
            Set-ADUser $samName -SamAccountName $new_sam -Verbose -ErrorAction Ignore                           

            }catch{
            Write-Warning $Error
            }
            }
function displayName{
[cmdletBinding()]
        param(

            [Parameter(Mandatory)]
            [String] $fname,
            [Parameter(Mandatory)]
            [String] $lname,
            [Parameter(Mandatory)]
            [String] $samName,
            [Parameter(Mandatory)]
            [String] $upn1
            )

            #Setting new email address
            $upn = "("+"$upn1"
            Set-ADUser $samName -DisplayName $lname", "$fname" "$upn -Verbose -ErrorAction Ignore
}

#UPN change
function change_email{
[cmdletBinding()]
        param(

            [Parameter(Mandatory)]
            [String] $oldemail,
            [Parameter(Mandatory)]
            [String] $newemail,
            [Parameter(Mandatory)]
            [String] $samName
            )

        try{
            #Adding attributes for new email and removing the old attributes
            $old_smtp = $oldemail.split("@")[0]+ ".cau@carnivalcorp.mail.onmicrosoft.com"
            $new_smtp = $newemail.split("@")[0]+ ".cau@carnivalcorp.mail.onmicrosoft.com"
            Set-ADUser $samName -remove @{proxyAddresses="SMTP:$oldemail"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -remove @{proxyAddresses="smtp:$old_smtp"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -remove @{proxyAddresses="sip:$oldemail"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -Add @{proxyAddresses="smtp:$new_smtp"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -Add @{proxyAddresses="sip:$newemail"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -Add @{proxyAddresses="SMTP:$newemail"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -Add @{proxyAddresses="smtp:$oldemail"} -Verbose -ErrorAction Ignore
            Set-ADUser $samName -EmailAddress $newemail -UserPrincipalName $newemail -Verbose -ErrorAction Ignore
            }catch{
            Write-Warning "Error from change_email function"
            }

}

#main function
function main{
$ask1 = 'y'
 while($ask1 -eq 'y' -or $ask1 -eq 'Y'){
        Write-Host "###################AD Console#################"
        Write-Host "**********************************************" -BackgroundColor Gray
        Write-Host "Create AD Account -option to copy gp press 1 " 
        Write-host "----------------------------------------------" -ForegroundColor DarkMagenta
        Write-Host "Name Change(FirstName or LastName)   press 2 "
        Write-host "----------------------------------------------" -ForegroundColor DarkMagenta
        Write-Host "Change Department(Email/Manager)     press 3 "
        Write-host "----------------------------------------------" -ForegroundColor DarkMagenta
        Write-Host "Termination                          press 4 "
        Write-Host "**********************************************" -BackgroundColor Gray      
        try{       
                $ask = Read-Host "Select 1-4, Q to quit." 
                switch($ask){

                #Create AD acccount
                "1"{
                Write-Host "-------------------Creating New AD----------------------" -BackgroundColor Green
                $userName= Read-host "Whose ID would you like to clone/copy?"
                $f_Name = Get-ADUser -Identity $userName -Properties *
                $fullName=$f_Name.Name
                Write-Host " You are copying $fullName AD info..."
                Write-Host "----------------------------------------------" -BackgroundColor DarkGray
                #getting replication user information that need to be copied
                $info = Get_Info -userID $userName
                [String]$email= $info[0]
                [String]$manager1= $info[1]
                [String]$location= $info[2]
                [String]$displayName= $info[3]
                #getting New user information that need to be added
                $full_d = user_details 
                [String] $first_N = $full_d[0] 
                [String] $last_N = $full_d[1]
                [String] $title = $full_d[2]
                [String] $dept_M = $full_d[3]
                #calling for Unique ID
                $samName= Create_UniqueID -fname $first_N -lname $last_N
                $final_email = "$first_N"+"."+"$last_N"+"@"+"$email"
                $final_displayName = "("+"$displayName"
                #creating AD           
                User_AD -firstName $first_N -lastName $last_N -samName $samName -title $title -Departname $dept_M -email $final_email -UPN $final_displayName -path $location -manager $manager1
                #Adding attributes to new AD
                default_attributes -sam_Name $samName -email $final_email
                #Adding Group Members to new AD 
                Write-Host $first_N" "$last_N  "AD has been created Successfully" -ForegroundColor Green                           
                Copy_group -FirstID $samName -SecondID $userName 
                default_groups -samname $samName
                $excel = 'y'#Read-Host "Would you like to create User Excel File in G:\ (Y/N) ?"
                if($excel -eq 'y' -or $excel -eq 'Y'){
                create_excel -firstName $first_N -lastName $last_N -samName $samName
                }
                else{
                Write-Host "AD process completed"
                }
                }

                "2"{
                #Chaning Name for e.g 1st name or last name
                Write-Host "--------------- Changing User Name----------------------" -BackgroundColor Green                              
                $userid = Read-Host "Enter User UserID"
                $fname = Read-Host "Enter user --(new)-- first name"
                $lname = Read-Host "Enter user --(new)-- last name"
                #changing 1st letter to capital and rest lower case
                $fixed_Names = Naming_case -first $fname -second $lname
                $f_name = $fixed_Names[0]
                $l_name = $fixed_Names[1]
                #calling function that hold the set attributes
                change_name -firstName $f_name -lastName $l_name -samName $userid
                Write-Host "---------------Successfully Completed-------------------" -BackgroundColor Green
                

                }
                "3"{
                #If someone department title and email changes
                Write-Host "--------------- Changing User Position email/Department----------------------" -BackgroundColor Green
                $userid = Read-Host "Enter UserID, want to copy access from"
                $user_change = Read-Host "Enter UserID, whose position you are about to change"
                $title = Read-Host "Enter User Job Title"
                $depM  = Read-Host "Enter User Department"
                $info = Get_Info -userID $userid
                [String]$email= $info[0]
                [String]$manager1= $info[1]
                [String]$location= $info[2]
                [String]$displayName= $info[3]
                #Accessing Properties
                $sammy = Get-ADUser -Identity $user_change -Properties *
                $first_N = $sammy.GivenName
                $last_N = $sammy.Surname

                $oldemail = $sammy.EmailAddress
                $new_email = "$first_N"+"."+"$last_N"+"@"+"$email"

                #changing email/Proxy and adding alies 
                change_email -oldemail $oldemail -newemail $new_email -samName $user_change
                #changing display name
                displayName -fname $first_N -lname $last_N -samName $user_change -upn1 $displayName

                #Changing properties
                Set-ADUser $user_change -Description $title -Department $depM -Title $title -Manager $manager1 -Verbose -ErrorAction Ignore
                copy_group -FirstID $user_change -SecondID $userid
                Write-Host "Move User to specific department into AD" -ForegroundColor Yellow
                

                Write-Host "--------------- Successfully Completed------------------" -BackgroundColor Green

   
                }

                "4"{
                Write-Host "--------------- You are about to Terminate user----------------------" -BackgroundColor Red
                ### Terminating user with removing access to drives and creating excel for the given user
                ### Reads input of username 
                $termuser = read-host "Enter userID" 
                $sammy = Get-ADUser -Identity $termuser -Properties *
                $first_N = $sammy.GivenName
                $last_N = $sammy.Surname

                $ask = Read-Host " You are about to TERMINATE $first_N" "$last_N Continye(Y),press anykey to quit?" 
                
                if($ask -eq 'Y' -or $ask -eq 'y'){
                
                
                     ### Set today's date for ADAccount expiration 
                    $termDate = get-date -uformat "%d/%m/%Y"
                     Set-ADAccountExpiration -Identity $termuser -DateTime $termDate -verbose
                     write-host "* $termuser account has been expired on $termDate *"

                     $DepartM = $sammy.Department
                     $manager =$sammy.Manager.Split("=")[1].split(",")[0]
                     $termObjectGUID = Get-ADUser $termuser -Properties objectGUID
                     Move-ADObject -identity $termObjectGUID -TargetPath 'OU=_Disabled,OU=Accounts,OU=_CAU,DC=carnivalaustralia,DC=com' -verbose
                     write-host "* $termuser moved to Disabled Users *"
                  ### Creating Excelfile.
                  ter_excel -firstName $first_N -lastName $last_N -depart $DepartM -Manager $manager -time $termDate -SamName $termuser
                   
                  ### Delete H drive
                     remove-item \\caunsfs01\users\$termuser -recurse -verbose -ErrorAction Ignore
                     Write-host "* $termuser H Drive is deleted *"
                  ### Delete Citrix profile 
         

                  ### Delete V drive
                     remove-item \\caunsms01\mailarchives\$termuser -recurse -verbose -ErrorAction Ignore
                     write-host "* $termuser V Drive is deleted *"

                  ### Clear Manager field in Organization 
                     Set-ADUser $termuser -clear Manager -verbose
                     write-host "* $termuser removed Manager field *"

                  ### Adds CAU to intial field 
                     Set-ADUser $termuser -Initials 'CAU' -verbose
                     write-host "* $termuser added 'CAU' to Initials *"

                  ### Change Description to "Terminated DD/MM/YYYY by CURRENT USER"
                      $terminatedby = $env:username
                      $termUserDesc = "Terminated " + $termDate + " by " + $terminatedby
                      set-ADUser $termuser -Description $termUserDesc 
                      write-host "* " $termuser "description set to" $termUserDesc

                  ### Disable user
                     Disable-ADAccount -Identity $termuser -verbose
                     write-host "*** TERMINATION COMPLETED" $termuser "account has been disabled ***"
                
                 
                } 

                else{
                Write-Host "Termination has been canceled" -ForegroundColor Red
                }

               
                }

               }

               if($ask -eq 'q' -or $ask -eq 'Q'){
               Write-host "Exited "
               break
               }
                }catch{
                Write-Error $Error[0]
                }
                $ask1= Read-Host "Would you like to Continue Again? Y/N ? "
                if($ask1 -eq 'N' -or $ask1 -eq 'n'){
                Write-Host "Good Bye"
  
                }

}
}
main