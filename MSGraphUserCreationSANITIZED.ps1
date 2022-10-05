#MUST BE RAN AS DOMAIN ADMIN FOR ACCOUNT CREATION
#version 1.1


$appid = 'AZURE APP ID HERE'
$tenantid = 'fakecompany.onmicrosoft.com'
$secret = 'AZURE APP SECRET HERE'
 
$body =  @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $appid
    Client_Secret = $secret
}
 
$connection = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token -Body $body
$token = $connection.access_token



Connect-MgGraph -AccessToken $token




$siteID = "SITE ID HERE" #the sharepoint site id 
$listId = "LIST ID HERE" #sharepoint list id here 
$listItemID = Get-MgSiteListitem -SiteId $siteID -ListId $listId 


foreach($entry in $listItemID){
    

    $ListItems = Get-MgSiteListitem -SiteId $siteID -ListId $listId -ListItemId $entry.id -ExpandProperty "fields" 
    $newUser = $ListItems.Fields.AdditionalProperties
    $givenName = $newUser.First #here First is the name of the sharepoint list field we're assigning if you named your field something else it will be that
    $surname = $newUser.Last #here Last is the name of the sharepoint list field we're assigning if you named your field something else it will be that

    if ($surname.Split("").count -gt 1){     #we're checking for spaces here and only keeping the second element if last name contains
        $surname = $newUser.Last.Split(" ")[1]
    }
    elseif ($surname.Split("-").count -gt 1){            #we're checking for hyphens here and only keeping the second element if last name contains
        $surname = $newUser.Last.Split("-")[1]
    }
    $site = $newUser.Site #here Site is the name of the sharepoint list field we're assigning if you named your field something else it will be that
    $title = $newUser.EmpTitle #here EmpTitle is the name of the sharepoint list field we're assigning if you named your field something else it will be that
    $empID = $newUser.EmployeeID #here EmployeeID is the name of the sharepoint list field we're assigning if you named your field something else it will be that
    $userDep = $newUser.Department #here Department is the name of the sharepoint list field we're assigning if you named your field something else it will be that
    $userName = $givenName.Substring(0,1)+$surname #how we do usernames is first initial last name so if you have a different method this will need to change
    $count = 2
    $path = $null
    While ([boolean] (Get-ADUser -Identity $userName)){ # check if username already exists only supports up to 9 ///Additional info here. This is kinda whack but it works good enough for us.
        #basically we check if the user exists and if it does we will append $count to the end and check again. If still exists we increment and reappend. Shitty solution but we never have that many people with same first initial last name. 
        
        $userName = $userName+$count.ToString() #here we add 2 to the end to create a unique username and check again 
        $count++ #incrementing count so that if username2 isnt unique it will remove and reappend up to 9 
        if([boolean] (Get-ADUser -Identity $userName)) { #check again if user is unique True here means not unique
            $userName = $userName.Substring(0, $userName.Length-1) #remove number then rerun loop 
        } 
        else {break}
        
        
    }
    

    write $path
    
    switch ($site) #switch statement to set users site specific info, this should really be different classes we inherit but this was written a while ago and just has never been updated
    {
        'SITE1' { #we set specific info based on users site. The switch statement is checking and will define what we want user attributes to be based on this value
            $path = 'OU=Site1,DC=FAKE,DC=LOCAL'
            $addr = '10 FAKE SITE, Ste 420'
            $city = 'FAKETOWN'
            $state = 'NE'
            $zip = 'ZIP'
            $phone = 'PHONE'
        }

        default {
            
        }
    }
    
    switch ($userDep) #switch statement to set users OU depending on department
    {
        'Administration' { #Our AD is structured department > site > fakecompany.local so here we're just creating the full path by appending to what we had assigned path to in the site switch statements
            $path = 'OU=Administration,'+$path #all this crap should probably be part of a class instead of a switch statement but what ever
            
            

        }

        'BackOffice' {
            $path = 'OU=Back Office,'+$path
            
        }
        

        'Billing' {
            $path = 'OU=Billing,'+$path
            
        }

        default { #if you're going to use switch statements instead of classes just add them before the default response
            
        }
    }


    
    
    #Create AD user - MUST BE RAN AS DOMAIN ADMIN BECAUSE OF THIS COMMAND
    New-ADUser -DisplayName "$givenName $surname" -Name $userName -UserPrincipalName $userName"@fakecompany.com" -office $site -Department $userDep -Title $title -GivenName $givenName -Surname $surname -Path $path -StreetAddress $addr -State $state -PostalCode $zip -EmailAddress $userName"@fakecompany.com" -OfficePhone $phone -EmployeeID $empID -City $city -AccountPassword (ConvertTo-SecureString "TempPassHere" -AsPlainText -Force)
    Enable-ADAccount -Identity $userName #enable account
    Add-ADGroupMember -Identity "Web Restricted" -Members $userName #add user to group
    $proxy = "SMTP:$($userName)@fakecompany.com","smtp:$($userName)@fakecompany.com","smtp:$($userName)@fakecompany.onmicrosoft.com" #define proxy address for fakecompany email addresses
    set-aduser -Identity $userName -add @{proxyaddresses = $proxy} #set proxy address for fakecompany email

    #remove entry from list
    if (get-aduser $userName) { #we now check to make sure the user was created and if so we will delete the object from the sharepoint list
        Remove-MgSiteListItem -SiteId $siteId -ListId $listId -ListItemId $entry.Id
    }
    
}

#END NOTES
#I havent looked at this script in a long time but it definitely could be done a lot better. I hope it at least helps others 
#in some way even if its not a perfect drop in solution for your work. If you have any questions shoot me an email dnbeze55@gmail.com