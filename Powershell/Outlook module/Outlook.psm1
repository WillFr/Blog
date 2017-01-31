

# Queries an outlook folder : return a list of items matchign the query
function Query-OutlookFolder($Folder, $Query)
{
    $r = @()
    foreach($f in $Folder)
    {
        $items = $f.Items
        $e = $items.Find($Query)

        while($e){
            $r += $e
            $e = $items.FindNext()
        }
    }
    return $r
}

# Attach to or start outlook
function Get-Outlook
{
    if(-not (Get-Process outlook))
    {
        $outlook = New-Object -comobject outlook.application
    }
    else
    {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    }
    return $outlook
}

# Return a list of folders matching the path. It support the wildcard '*'
# the path separator can be either '/' or '\'
# eg: Inbox/*
function Get-OutlookFolder($Outlook, $Path)
{
    $inbox = $Outlook.GetNameSpace("MAPI").GetDefaultFolder(6)

    $arr = $Path.Split(@('/','\'))
    $i = 0
    if($arr[0] -ieq "Inbox")
    {
        $cur = $inbox
        $i = 1
    }
    else 
    {
        $cur =$inbox.Parent
    }

    
    while($i -ne $arr.Length -and $arr[$i] -ne "*")
    {
        $cur = $cur.Folders | where { ($_.FolderPath.Split(@('/','\'))  | Select-Object -Last 1 ) -eq $arr[$i] }
        $i++
    }

    if($i -ne $arr.Length -and $arr[$i] -eq "*")
    {
        $r =@()
        $q = New-Object System.Collections.Queue
        $q.Enqueue($cur)

        while($q.Count -ne 0){
            $cur = $q.Dequeue()
            $r += $cur
            foreach($folder in $cur.Folders | where { $_.DefaultMessageClass -eq "IPM.Note"})
            {
                $q.Enqueue($folder)
            }
        }

        return $r
    }
    else{
        return $cur
    }
}


# starts a webserver to open emails from an html request
function Start-OutlookEmailLinkServer($LocalEndpoint, $Folders, [switch]$InProcess)
{
    if($InProcess)
    {
        Write-Host "Initializing outlook server on folders : "
        $Folders | foreach { Write-Host $_ }
        $outlook = Get-Outlook
        $outlookFolders = $Folders | foreach { Get-OutlookFolder -Outlook $outlook -Path $_ }
        $server = New-Object Net.HttpListener
        $server.Prefixes.Add($LocalEndpoint)
        $server.Start()
        Write-Host "Launching server : if your PS window is blocked, visit $LocalEndpoint in a browser"
        [Environment]::SetEnvironmentVariable("OutlookServerEndpoint", $LocalEndpoint, "User")
        try
        {
            do
            {
        
                $Context = $server.GetContext()
                Write-Output "Request received $($Context.Request.RawUrl)"
                Write-Output "`tif your PS window is blocked, visit $LocalEndpoint in a browser"
                if($Context.Request.RawUrl -eq "/") 
                {
                    Write-Output "Exit request received $($Context.Request.RawUrl) : stopping outlook server"
                    break
                }
                if($Context.Request.RawUrl -eq "/a") 
                {
                    continue
                }

                $emailId = $Context.Request.RawUrl.Substring(2)
                $email = $outlook.GetNameSpace("MAPI").GetItemFromID($emailId)

                Write-Host "Looking for email about $($email.ConversationTopic) in $($email.Parent.FolderPath)"
                #$outlookFolders  | foreach { Write-Host $_.FolderPath }

                $lastEmail = Query-OutlookFolder -Folder $email.Parent -Query "[Conversation]=`"$($email.ConversationTopic)`"" | where { $_.MessageClass -ieq "IPM.Note" } | Sort -Property ReceivedTime | Select -Last 1
                $lastEmail.Display()
            }while($true)
        } 
        finally
        {
            $server.Stop()
        }

        #return
        exit
    }

    $serverEndpoint = [Environment]::GetEnvironmentVariable("OutlookServerEndpoint","User")

    try
    {
        # kill potentially running servers
        Invoke-RestMethod -Method GET -Uri $serverEndpoint -UseBasicParsing
        Invoke-RestMethod -Method GET -Uri $LocalEndpoint -UseBasicParsing
    }
    catch{}

    $modulePath = (Get-Module Outlook*).path
    $folderString = "@('$($Folders -join "','")')"
    Start-Process powershell `
        -ArgumentList "-NoExit","-Command `"& { ipmo $modulePath -force; Start-OutlookEmailLinkServer -LocalEndpoint $LocalEndpoint -Folders $folderString -InProcess}`"" #-WindowStyle Hidden
    #Start-OutlookEmailLinkServer -LocalEndpoint $LocalEndpoint -Folders $Folders -InProcess
}