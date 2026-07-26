# LibraryLint - HTPC Wake Control
# -------------------------------
# Wake-on-LAN + Kodi JSON-RPC for the HTPC that consumes LibraryLint's output.
# Originally a Dell OptiPlex 3060 running LibreELEC/Kodi, but works against
# any host where you know the MAC, IP, and a TCP port to use as the "up"
# signal (Kodi web port 8080, SSH 22, SMB 445 — whatever the host listens on).
#
# All functions take explicit parameters rather than reading $script:Config —
# the module loads via Import-Module which gives it its own script scope, so
# the main LibraryLint.ps1 must pass values in.
#
# Public functions:
#   Send-Wol           Broadcast a magic packet to any MAC.
#   Wait-Htpc          Poll a TCP port until it answers or a timeout elapses.
#   Start-Htpc         Send WoL; optionally block until the signal port answers.
#   Invoke-KodiJsonRpc POST a single JSON-RPC method to Kodi's web port.
#   Stop-Htpc          Ask Kodi to shutdown / suspend / hibernate / reboot.

<#
.SYNOPSIS
    Broadcasts a Wake-on-LAN magic packet to a MAC address.
.DESCRIPTION
    Builds the standard 102-byte magic packet (6 bytes of 0xFF followed by
    the target MAC repeated 16 times) and broadcasts it as a UDP datagram.
    Defaults to limited broadcast (255.255.255.255) on port 9, but subnet-
    directed broadcast (e.g. 192.168.0.255) is more reliable on switches
    that drop limited broadcasts.
.PARAMETER Mac
    Target MAC address. Accepts colon- or dash-separated hex pairs.
.PARAMETER Broadcast
    Broadcast address to send to. Prefer the subnet-directed form
    (192.168.0.255) over 255.255.255.255 — some gear drops the latter.
.PARAMETER Port
    UDP port. WoL is port-agnostic on the receiving NIC, but convention is
    7 (echo) or 9 (discard). 9 is the safer default.
.EXAMPLE
    Send-Wol -Mac '6c:2b:59:db:eb:5b'
.EXAMPLE
    Send-Wol -Mac '6c:2b:59:db:eb:5b' -Broadcast '192.168.0.255'
#>
function Send-Wol {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$')]
        [string]$Mac,

        [string]$Broadcast = '255.255.255.255',
        [int]$Port = 9
    )

    $macBytes = $Mac -split '[:\-]' | ForEach-Object { [byte][Convert]::ToInt32($_, 16) }
    $packet   = [byte[]]( @(0xFF) * 6 + ($macBytes * 16) )   # 6x 0xFF + MAC x16 = 102 bytes

    $udp = [System.Net.Sockets.UdpClient]::new()
    $udp.EnableBroadcast = $true
    try {
        $udp.Connect($Broadcast, $Port)
        [void]$udp.Send($packet, $packet.Length)
        Write-Verbose "Magic packet sent to $Mac via ${Broadcast}:${Port}"
    }
    finally {
        $udp.Dispose()
    }
}

<#
.SYNOPSIS
    Polls a TCP port on the HTPC until it answers or the timeout elapses.
.DESCRIPTION
    Uses Test-Connection -TcpPort (PowerShell 7+) which probes the actual
    TCP handshake rather than ICMP. Kodi's web port (8080 by default) is a
    good "Kodi is up and serving" signal; SSH (22) is a good "OS booted"
    signal even on Kodi-less builds.
.PARAMETER IPAddress
    Target host IP. Required.
.PARAMETER Port
    TCP port to probe (default 8080 = Kodi web/JSON-RPC).
.PARAMETER TimeoutSeconds
    Total wait budget. Boot from a cold OptiPlex 3060 to Kodi web ready is
    typically 30-60s; the 90s default leaves headroom.
.PARAMETER IntervalSeconds
    Pause between probe attempts.
.OUTPUTS
    [bool] - $true if reachable before timeout, else $false.
#>
function Wait-Htpc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$IPAddress,
        [int]$Port            = 8080,
        [int]$TimeoutSeconds  = 90,
        [int]$IntervalSeconds = 3
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-Connection -TargetName $IPAddress -TcpPort $Port -Quiet) {
            Write-Verbose "${IPAddress}:${Port} reachable."
            return $true
        }
        Start-Sleep -Seconds $IntervalSeconds
    }
    Write-Warning "${IPAddress}:${Port} not reachable within ${TimeoutSeconds}s."
    return $false
}

<#
.SYNOPSIS
    Sends a WoL magic packet and (optionally) waits for the host to answer.
.PARAMETER Mac
    Target MAC address (required).
.PARAMETER Broadcast
    Broadcast address for the magic packet. Subnet-directed (192.168.0.255)
    is more reliable than 255.255.255.255 on some switches.
.PARAMETER IPAddress
    Required when -Wait is set; used to poll the readiness port.
.PARAMETER Port
    TCP signal port for the readiness probe (default 8080 = Kodi web).
.PARAMETER Wait
    Block after sending the packet until the readiness port answers or the
    timeout elapses.
.PARAMETER TimeoutSeconds
    Maximum wait budget when -Wait is set.
.OUTPUTS
    [bool] when -Wait is used (reachable / not). Nothing otherwise.
.EXAMPLE
    Start-Htpc -Mac '6c:2b:59:db:eb:5b' -Broadcast '192.168.0.255'
.EXAMPLE
    if (Start-Htpc -Mac '6c:..' -Broadcast '..' -IPAddress '192.168.0.191' -Wait) {
        Write-Host "HTPC is up."
    }
#>
function Start-Htpc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Mac,
        [string]$Broadcast = '255.255.255.255',
        [string]$IPAddress,
        [int]$Port         = 8080,
        [switch]$Wait,
        [int]$TimeoutSeconds = 90
    )

    Send-Wol -Mac $Mac -Broadcast $Broadcast
    Write-Verbose "Wake sent to HTPC ($Mac via $Broadcast)."

    if ($Wait) {
        if (-not $IPAddress) {
            Write-Warning "-Wait requires -IPAddress."
            return $false
        }
        return Wait-Htpc -IPAddress $IPAddress -Port $Port -TimeoutSeconds $TimeoutSeconds
    }
}

<#
.SYNOPSIS
    POSTs a single Kodi JSON-RPC call to the HTPC's web interface.
.DESCRIPTION
    Kodi exposes its full RPC surface at /jsonrpc on its web port (default
    8080). Low-level helper used by Stop-Htpc and any future Kodi-driven
    feature (library scan trigger, currently-playing info, etc.).

    Auth: if -User and -Password are both supplied, sends them as HTTP
    Basic. Most LibreELEC builds ship with web auth disabled by default —
    leave them off in that case.
.PARAMETER IPAddress
    HTPC IP (required).
.PARAMETER Port
    Kodi web port (default 8080).
.PARAMETER Method
    Kodi JSON-RPC method name (e.g., "System.Shutdown", "System.Suspend").
.PARAMETER Params
    Optional params object for the call. Some methods take none.
.PARAMETER User
    HTTP Basic username (optional). Only sent if both User and Password set.
.PARAMETER Password
    HTTP Basic password (optional). Only sent if both User and Password set.
.PARAMETER TimeoutSeconds
    HTTP timeout. Shutdown/suspend calls return immediately; a sleeping
    HTPC's port is unreachable — keep short so a missing box doesn't hang.
.OUTPUTS
    Hashtable: Success (bool), Result (parsed JSON-RPC result on success),
    Error (string on failure).
#>
function Invoke-KodiJsonRpc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$IPAddress,
        [int]$Port = 8080,
        [Parameter(Mandatory)] [string]$Method,
        [object]$Params,
        [string]$User,
        [string]$Password,
        [int]$TimeoutSeconds = 8
    )

    $body = @{
        jsonrpc = '2.0'
        method  = $Method
        id      = 1
    }
    if ($PSBoundParameters.ContainsKey('Params')) {
        $body.params = $Params
    }
    $json = $body | ConvertTo-Json -Compress -Depth 10

    $uri = "http://${IPAddress}:${Port}/jsonrpc"
    $invokeParams = @{
        Uri         = $uri
        Method      = 'Post'
        Body        = $json
        ContentType = 'application/json'
        TimeoutSec  = $TimeoutSeconds
        ErrorAction = 'Stop'
    }
    if ($User -and $Password) {
        $secure = ConvertTo-SecureString $Password -AsPlainText -Force
        $invokeParams.Credential = [PSCredential]::new($User, $secure)
    }

    try {
        $response = Invoke-RestMethod @invokeParams
        if ($response.error) {
            return @{ Success = $false; Error = "$($response.error.message) (code $($response.error.code))" }
        }
        return @{ Success = $true; Result = $response.result }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Asks the HTPC to shutdown, suspend, hibernate, or reboot via Kodi JSON-RPC.
.DESCRIPTION
    Issues a System.<mode> call. The HTPC stops responding within a few
    seconds — once it sleeps, the port goes away — so the call returning
    success means Kodi received and acted on the request.

    Default mode is Shutdown (S5) — pairs with the cold-boot WoL flow.
    Suspend (S3) wakes faster but needs BIOS support for resume-on-magic-
    packet from S3. Hibernate (S4) is rarely worth using on an HTPC.
.PARAMETER IPAddress
    HTPC IP (required).
.PARAMETER Port
    Kodi web port (default 8080).
.PARAMETER Mode
    Shutdown (default) | Suspend | Hibernate | Reboot.
.PARAMETER User / Password
    Optional HTTP Basic credentials for Kodi web.
.OUTPUTS
    Hashtable: Success (bool), Error (string on failure).
#>
function Stop-Htpc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$IPAddress,
        [int]$Port = 8080,
        [ValidateSet('Shutdown', 'Suspend', 'Hibernate', 'Reboot')]
        [string]$Mode = 'Shutdown',
        [string]$User,
        [string]$Password
    )

    $method = "System.$Mode"
    $rpcParams = @{
        IPAddress = $IPAddress
        Port      = $Port
        Method    = $method
    }
    if ($User)     { $rpcParams.User     = $User }
    if ($Password) { $rpcParams.Password = $Password }

    $result = Invoke-KodiJsonRpc @rpcParams

    if ($result.Success) {
        Write-Verbose "$Mode request acknowledged by Kodi at $IPAddress."
    } else {
        Write-Verbose "$Mode request failed: $($result.Error)"
    }
    return $result
}

Export-ModuleMember -Function Send-Wol, Wait-Htpc, Start-Htpc, Stop-Htpc, Invoke-KodiJsonRpc
