#In this script, we will be connecting to M365 tenant and createing new groups. 
#Everything will be logged in a notepad file.
#Author: Dharmik Pandya
#Date: July 17, 2022

Set-ExecutionPolicy RemoteSigned


$menuInput = 0                                              #Default Menu option 
$todaysDate = Get-Date -Format "MM-dd-yyyy-hh-mm"           #Date format for log file
$groupName = $null
#File log time stamp format
#Source: https://www.gngrninja.com/script-ninja/2016/2/12/powershell-quick-tip-simple-logging-with-timestamps
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}


$logFile = "C:\Users\Public\logFile-" + $todaysDate + ".txt"    #Path of the file location
New-Item -Path $logFile -ItemType File | Out-Null               #Creating new log file each time the script runs

#first confirmaton of log creation
Write-Output "$(Get-TimeStamp) Log file created" | Out-file $logFile -append

<#
Starting with Try and catch block. If the login is successful, it will proceed with the script, if not, it will
end the script.
#>
try {

    #logging
    Write-Output "$(Get-TimeStamp) User prompted to sign in to MS365 tenant" | Out-file $logFile -append

    #Disclaimer for the user
    Write-Output "#Welcome, this script allows you to add a group using powershell commands. 
    Please note, from the moment you begin this script, logging of each input is enabled.
    Script logs will be stored in C:\Public\Public Documents\ location with only read rights."

    Start-Sleep -Seconds 2                                      #haulting script

    Write-Output "`nPlease sign in to AzureAD, if you do not see a prompt, minimize the powershell window`n"

    Start-Sleep -Seconds 2                                      #haulting script

    #Connecting to AzureAD
    Connect-AzureAD -ErrorAction SilentlyContinue

    <#
    Running if else block, if the Account is authorized, it will proceed with the script.
    #>
        
    #logging starts for authentication
    Write-Output "$(Get-TimeStamp) User logged in successfully" | Out-file $logFile -append
    Write-Output "Conncetion is successful to $(Get-AzureADTenantDetail | Select-object -Property DisplayName)"
    Write-Output "*********************** || *************************" | Out-file $logFile -append
    Write-Output "$(Get-TimeStamp) Connection was successful to M365 Tenant" | Out-file $logFile -append
    Write-Output "Connected to $((Get-AzureADTenantDetail).DisplayName)" | Out-file $logFile -append
    Write-Output "Connected to $((Get-AzureADTenantDetail).ObjectID)" | Out-file $logFile -append
    Write-Output "*********************** || *************************" | Out-file $logFile -append
    #logging ends for authentication

    #Menu for the users to choose from
    $menu = "

        Please choose an option from the menu below

        1. Add a group
        2. Add an owner to a group
        3. View session logs
        4. Exit
        "
    <#Running While loop for user to choose between the different menu options. Loop breaker is value of Pie#>
    while ($menuInput -ne 3.14) {
        Write-Output $menu                                                  #Printing out menu

        $menuInput = Read-Host -prompt 'Choose an option above'             #Prompting user to choose

        #Option 1: Adding a group to AzureAD
        if ($menuInput -eq 1) {

            #Logging
            Write-Output "$(Get-TimeStamp) User promoted to add a group" | Out-file $logFile -append  

            #informational
            Write-Output "`nThis script is only capable of creating a Security group only`n"

            #prompting user to enter the name of the group
            $groupName = Read-Host -prompt 'Enter the name of the group: '

            #prompting user to enter the description of the group
            $groupDescription = Read-Host -prompt 'Enter the description of the group: '

            #Logging
            Write-Output "$(Get-TimeStamp) Group to add: $groupName" | Out-file $logFile -append 
            Write-Output "$(Get-TimeStamp) Description to add: $groupDescription" | Out-file $logFile -append 

            #Checking if the group already exists
            if ((Get-AzureADGroup).DisplayName -eq $groupName) {
                Write-Host "`n******************E*R*R*O*R***********************`n"
                Write-Host "Group, $groupName, already exists. Please try again"
                Write-Host "`n******************E*R*R*O*R***********************`n"

                $groupName = $null                                          #resetting the variable

                #Logging
                Write-Output "$(Get-TimeStamp) ERROR Group already exists: $groupName" | Out-file $logFile -append 
            }
            #Add the group if it does not exist
            else {
                #Adding group to AzureAD
                New-AzureADGroup -DisplayName $groupName -Description $groupDescriptio -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
                    
                #Printing a note
                Write-Host "`nCreating group in AzureAD... Please Standby`n`n"
                Start-Sleep -Seconds 3                                      #Haulting script
                Write-Host "`n`n*******Group is created*******"             #Printing a note
                Write-Host "`n`n`n"

                #Logging
                Write-Output "$(Get-TimeStamp) Group, $groupName, has been added." | Out-file $logFile -append 
            }
        }

            
        #Option 2: Adding an owner to an existing group or to the above group
        elseif ($menuInput -eq 2) {
                
            #Logging
            Write-Output "$(Get-TimeStamp) User has prompted to add owner to a group" | Out-file $logFile -append

            #Printing disclaimer
            Write-Host "`nYou have selected an option to add owner to a group`n"

            $extGroupName = Read-Host -Prompt "Enter group name "          #Prompting user to add a group

            #Checking if the group exists in the AzureAD Tenant
            if ((Get-AzureADGroup).DisplayName -eq $extGroupName) {
                #Prompting user to add a user's email
                $gpOwnerEmail = Read-Host -Prompt "Enter user's email "

                #Checking if the email entered by the user is valid in Azure Tenant
                if ((Get-AzureADUser).UserPrincipalName -eq $gpOwnerEmail) {
                    $extgrpObjID = $(Get-AzureADGroup -Filter "DisplayName eq '$extGroupName'").ObjectId                   #getting Obecjt ID for the Group
                    $extuserObjID = $(Get-AzureADUser -Filter "UserPrincipalName eq '$gpOwnerEmail'").ObjectId                 #Getting Object ID for the user
                    $extGroupOwners = (Get-AzureADGroupOwner -ObjectId $extgrpObjID | Select ObjectId)

                    #Checking if the email entered by the user is already an owner of the group
                    if ($extGroupOwners -Match $extuserObjID) {

                        #Logging
                        Write-Output "$(Get-TimeStamp) User, $gpOwnerEmail, is already owner of, $extGroupName, group" | Out-file $logFile -append 
                        
                        #Printing result
                        Write-Host "`n`nUser, $gpOwnerEmail, is already an owner of the group, $extGroupName.`n`n "
                    }
                    else {
                        #Adding owner to the requested group
                        Add-AzureADGroupOwner -ObjectId $extgrpObjID -RefObjectId $extuserObjID
                            
                        #Logging
                        Write-Output "$(Get-TimeStamp) User, $gpOwnerEmail, is added as the owner of the, $extGroupName, group" | Out-file $logFile -append 
                        
                        #Printing result
                        Write-Host "`n`nUser, $gpOwnerEmail, is added as the owner of the, $extGroupName, group`n`n"
                    }
                }
                else {
                    #Throwing error since group does not exist
                    Write-Host "`n******************E*R*R*O*R***********************`n"
                    Write-Host "User, $gpOwnerEmail, does not exist. Please enter a user from your Azure tenant."
                    Write-Host "`n******************E*R*R*O*R***********************`n"
        
                    #Logging
                    Write-Output "$(Get-TimeStamp) ERROR User does not exist: $gpOwnerEmail" | Out-file $logFile -append
                }
                       
                    
            }
            #Throwing error since group does not exist
            else {
                Start-Sleep -Seconds 1                                      #haulting Script

                Write-Host "`n******************E*R*R*O*R***********************`n"
                Write-Host "Group, $extGroupName, does not exist. Please enter the group from your Azure tenant."
                Write-Host "`n******************E*R*R*O*R***********************`n"
        
                #Logging
                Write-Output "$(Get-TimeStamp) ERROR Group does not exist: $extGroupName" | Out-file $logFile -append
            }
                
        }
            
        #Option 3: Printing out log life
        elseif ($menuInput -eq 3) {
            #logging
            Write-Output "$(Get-TimeStamp) User viewed logs" | Out-file $logFile -append
            Get-Content $logFile                                            #Printing logfile
        }

        #Option 4: Exiting script
        elseif ($menuInput -eq 4) {

            Write-Output "`n`nExiting... "

            Disconnect-AzureAD                                              #Disconnection session from AzureAD

            #Logging start for ending session
            Write-Output "$(Get-TimeStamp) User selected to end the script" | Out-file $logFile -append  
            Write-Output "$(Get-TimeStamp) END OF LOG" | Out-file $logFile -append  
            #logging ends for ending session

            Start-Sleep -Seconds 1.5                                         #Haluting Script
            Write-Output "`n`nScript has ended..."
            $menuInput = 3.14                                                #Loop breaker
        }
        #Throwing error for slecting an incorrect option
        elseif ($menuInput -le 0 -or $menuInput -ge 5) {
            Write-Output "$(Get-TimeStamp) User inputted invalid menu option: $menuInput" | Out-file $logFile -append  

            Write-Host "Invalid option... Please try again."
        }
    }
}
catch {

    #Write-Output "Ran into an issue: $($PSItem.ToString())" #This is a for debugging

    #logging starts for authentication error 
    Write-Output "Error in authentication... Please run the script again..."
    Write-Output "*********************** || *************************" | Out-file $logFile -append
    Write-Output "$(Get-TimeStamp) Connection was failed to M365 Tenant\n END OF LOG" | Out-file $logFile -append
    Write-Output "*********************** || *************************" | Out-file $logFile -append
    #logging ends for authentication error 
}

