<#
.SYNOPSIS
    Generates comprehensive reports and visualizations for L-Kennung user accounts and group memberships.

.DESCRIPTION
    This script consolidates L-Kennung_Report.ps1 and L-Kennung_Tabelle.ps1. It queries AD for users matching 'L*',
    processes their group memberships, and exports the data to CSV, Excel (with group color-coding), and HTML (interactive D3.js visualization).
    Optionally, if the respective applications are installed, it can export to Microsoft Visio and MindManager.

.PARAMETER SearchPattern
    The SAMAccountName pattern to search for (e.g. 'L*'). Defaults to 'L*'.

.PARAMETER OutputPath
    The path to save the CSV report. Defaults to 'C:\daten\AD_Benutzer_Gruppen_L.csv'.

.PARAMETER ExcelPath
    The path to save the Excel sheet. Defaults to 'C:\daten\AD_Benutzer_Gruppen_L.xlsx'.

.PARAMETER HtmlPath
    The path to save the interactive HTML visualization. Defaults to 'C:\daten\AD_Benutzer_Gruppen_L.html'.

.PARAMETER VisioPath
    The path to save the Visio diagram. If specified, the script attempts a Visio export.

.PARAMETER MindManagerPath
    The path to save the MindManager map. If specified, the script attempts a MindManager export.

.PARAMETER Force
    If specified, closes running Excel processes automatically before writing.

.NOTES
    Version: 1.0
    Autor: Systems Administration
    Compatibility: PowerShell 5.1+
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$SearchPattern = "L*",

    [Parameter()]
    [string]$OutputPath = "C:\daten\AD_Benutzer_Gruppen_L.csv",

    [Parameter()]
    [string]$ExcelPath = "C:\daten\AD_Benutzer_Gruppen_L.xlsx",

    [Parameter()]
    [string]$HtmlPath = "C:\daten\AD_Benutzer_Gruppen_L.html",

    [Parameter()]
    [string]$VisioPath = "",

    [Parameter()]
    [string]$MindManagerPath = "",

    [Parameter()]
    [switch]$Force
)

# Load environment configuration if available in script directory or parent directories
$envLoader = Join-Path $PSScriptRoot "..\..\Import-Environment.ps1"
if (Test-Path $envLoader) {
    . $envLoader
}

# Helper to close running Excel processes
function Close-ExcelProcesses {
    Get-Process -Name "excel" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $_.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (-not $_.HasExited) { $_.Kill() }
        } catch {
            Write-Warning "Could not close Excel process: $_"
        }
    }
    Start-Sleep -Seconds 1
}

# Queries AD and compiles structured user-group mapping
function Get-ADUserGroupData {
    param ([string]$Pattern)

    Write-Verbose "Querying AD users matching pattern: $Pattern"
    $users = Get-ADUser -Filter "SamAccountName -like '$Pattern'" -Properties SamAccountName, Name, MemberOf, DistinguishedName, Comment -ErrorAction SilentlyContinue
    
    if (-not $users) {
        Write-Warning "No users found matching pattern '$Pattern'."
        return $null
    }
    
    $userData = [System.Collections.Generic.List[PSObject]]::new()
    $groupColors = @{}
    $colorIndex = 35 # Starting index for Excel interior colors
    
    foreach ($user in $users) {
        # Determine OU Name and sorting prefix
        $ouMatch = $user.DistinguishedName -match 'OU=([^,]+)'
        $ou = if ($ouMatch) { $Matches[1] } else { "No OU" }
        $numericPrefix = if ($ou -match '^\d{2}') { $Matches[0] } else { "999" }
        
        # Get group names
        $groups = $user.MemberOf | ForEach-Object {
            try {
                (Get-ADGroup $_).Name
            } catch {
                Write-Warning "Could not resolve group DN '$_' for user $($user.SamAccountName)"
                return "Unknown Group"
            }
        } | Sort-Object
        
        if (-not $groups) {
            # User has no groups
            $userData.Add([PSCustomObject]@{
                SortPrefix = $numericPrefix
                OU = $ou
                UserName = $user.Name
                SamAccountName = $user.SamAccountName
                Group = "No Groups"
                Comment = $user.Comment
                ColorIndex = 0
            })
            continue
        }
        
        foreach ($group in $groups) {
            # Assign color index for excel coloring
            if (-not $groupColors.ContainsKey($group)) {
                $groupColors[$group] = $colorIndex
                $colorIndex++
                if ($colorIndex -gt 46) { $colorIndex = 35 } # Loop color index
            }
            
            $userData.Add([PSCustomObject]@{
                SortPrefix = $numericPrefix
                OU = $ou
                UserName = $user.Name
                SamAccountName = $user.SamAccountName
                Group = $group
                Comment = $user.Comment
                ColorIndex = $groupColors[$group]
            })
        }
    }
    
    return @{
        Data = $userData
        Colors = $groupColors
    }
}

function Export-ToCSV {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    $directory = Split-Path -Parent $Path
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    
    if (Test-Path $Path) {
        Remove-Item $Path -Force -ErrorAction Stop
    }
    
    $Data | Export-Csv -Path $Path -NoTypeInformation -Delimiter ";" -Encoding UTF8 -Force
    Write-Host "CSV export completed: $Path" -ForegroundColor Green
}

function Export-ToExcel {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][hashtable]$GroupColors,
        [Parameter(Mandatory)][string]$Path
    )
    
    $excel = $null
    try {
        $directory = Split-Path -Parent $Path
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        if (Test-Path $Path) {
            Remove-Item $Path -Force -ErrorAction Stop
        }
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        $headers = @("OU", "Benutzer", "SamAccountName", "Gruppe", "Kommentar")
        1..5 | ForEach-Object { 
            $worksheet.Cells.Item(1, $_) = $headers[$_ - 1]
        }
        
        $row = 2
        foreach ($item in $Data) {
            $worksheet.Cells.Item($row, 1) = $item.OU
            $worksheet.Cells.Item($row, 2) = $item.UserName
            $worksheet.Cells.Item($row, 3) = $item.SamAccountName
            $worksheet.Cells.Item($row, 4) = $item.Group
            $worksheet.Cells.Item($row, 5) = $item.Comment
            
            if ($item.Group -ne "No Groups") {
                $groupCell = $worksheet.Cells.Item($row, 4)
                $groupCell.Interior.ColorIndex = $GroupColors[$item.Group]
            }
            $row++
        }
        
        $headerRange = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells(1, 5))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($row - 1, 5)).AutoFilter() | Out-Null
        $worksheet.Columns.Item(1).ColumnWidth = 25
        $worksheet.Columns.Item(2).ColumnWidth = 30
        $worksheet.Columns.Item(3).ColumnWidth = 20
        $worksheet.Columns.Item(4).ColumnWidth = 45
        $worksheet.Columns.Item(5).ColumnWidth = 40
        
        $workbook.SaveAs($Path)
        $workbook.Close($true)
        Write-Host "Excel export completed: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "Excel export failed: $_"
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

function Export-ToHtml {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    $directory = Split-Path -Parent $Path
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    
    $jsonData = $Data | ConvertTo-Json
    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>AD Benutzer & Gruppen Visualisierung</title>
    <script src="https://d3js.org/d3.v7.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f8f9fa; }
        h1 { color: #343a40; text-align: center; }
        .node circle { stroke: #fff; stroke-width: 1.5px; }
        .node text { font: 11px sans-serif; pointer-events: none; }
        .link { fill: none; stroke: #cbd5e0; stroke-opacity: 0.6; stroke-width: 1px; }
        #visualization { width: 100%; height: 850px; background-color: #ffffff; border: 1px solid #dee2e6; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        .tooltip { position: absolute; padding: 8px; background: rgba(0, 0, 0, 0.8); color: #fff; border-radius: 4px; pointer-events: none; font-size: 12px; }
    </style>
</head>
<body>
    <h1>Active Directory Benutzer- und Gruppenzugehörigkeiten</h1>
    <div id="visualization"></div>
    <script>
        const data = { nodes: [], links: [] };
        const users = $jsonData;
        const processedUsers = new Set();
        const processedGroups = new Set();
        
        users.forEach(user => {
            if (!processedUsers.has(user.SamAccountName)) {
                data.nodes.push({
                    id: user.SamAccountName,
                    type: 'user',
                    name: user.UserName,
                    ou: user.OU
                });
                processedUsers.add(user.SamAccountName);
            }
            
            if (user.Group && user.Group !== "No Groups") {
                if (!processedGroups.has(user.Group)) {
                    data.nodes.push({
                        id: user.Group,
                        type: 'group',
                        name: user.Group
                    });
                    processedGroups.add(user.Group);
                }
                
                data.links.push({
                    source: user.SamAccountName,
                    target: user.Group
                });
            }
        });
        
        const width = document.getElementById('visualization').clientWidth;
        const height = 850;
        
        const svg = d3.select("#visualization")
            .append("svg")
            .attr("width", width)
            .attr("height", height);

        const container = svg.append("g");
        
        // Add zoom behavior
        svg.call(d3.zoom().on("zoom", (event) => {
            container.attr("transform", event.transform);
        }));
        
        const simulation = d3.forceSimulation(data.nodes)
            .force("link", d3.forceLink(data.links).id(d => d.id).distance(60))
            .force("charge", d3.forceManyBody().strength(-150))
            .force("center", d3.forceCenter(width / 2, height / 2))
            .force("collision", d3.forceCollide().radius(15));
        
        const link = container.append("g")
            .selectAll("line")
            .data(data.links)
            .join("line")
            .attr("class", "link");
        
        const node = container.append("g")
            .selectAll("g")
            .data(data.nodes)
            .join("g")
            .attr("class", "node")
            .call(d3.drag()
                .on("start", dragstarted)
                .on("drag", dragged)
                .on("end", dragended));
        
        node.append("circle")
            .attr("r", d => d.type === 'user' ? 6 : 9)
            .style("fill", d => d.type === 'user' ? "#4a90e2" : "#e06666");
        
        node.append("text")
            .text(d => d.name)
            .attr("x", 12)
            .attr("y", 3)
            .style("fill", "#4a5568");
        
        // Tooltip
        const tooltip = d3.select("body").append("div").attr("class", "tooltip").style("opacity", 0);
        
        node.on("mouseover", (event, d) => {
            tooltip.transition().duration(200).style("opacity", .9);
            let content = d.type === 'user' ? `<b>Benutzer:</b> \${d.name}<br/><b>SAM:</b> \${d.id}<br/><b>OU:</b> \${d.ou}` : `<b>Gruppe:</b> \${d.name}`;
            tooltip.html(content)
                .style("left", (event.pageX + 10) + "px")
                .style("top", (event.pageY - 28) + "px");
        }).on("mouseout", () => {
            tooltip.transition().duration(500).style("opacity", 0);
        });

        simulation.on("tick", () => {
            link
                .attr("x1", d => d.source.x)
                .attr("y1", d => d.source.y)
                .attr("x2", d => d.target.x)
                .attr("y2", d => d.target.y);
            
            node
                .attr("transform", d => `translate(${d.x},${d.y})`);
        });
        
        function dragstarted(event, d) {
            if (!event.active) simulation.alphaTarget(0.3).restart();
            d.fx = d.x;
            d.fy = d.y;
        }
        
        function dragged(event, d) {
            d.fx = event.x;
            d.fy = event.y;
        }
        
        function dragended(event, d) {
            if (!event.active) simulation.alphaTarget(0);
            d.fx = null;
            d.fy = null;
        }
    </script>
</body>
</html>
"@
    $html | Out-File -FilePath $Path -Encoding UTF8
    Write-Host "HTML visualization completed: $Path" -ForegroundColor Green
}

function Export-ToVisio {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    $visio = $null
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false
        
        $doc = $visio.Documents.Add("")
        $page = $doc.Pages.Item(1)
        
        # Open standard organization chart stencil
        $stencil = $visio.Documents.OpenEx("ORGCH_M.VSS", 4)
        $userShape = $stencil.Masters.Item("Process")
        $groupShape = $stencil.Masters.Item("Decision")
        
        $shapes = @{}
        $yPos = 1
        $xPos = 1
        
        # Render Users
        $Data | Sort-Object OU, UserName -Unique | ForEach-Object {
            if (-not $shapes.ContainsKey($_.SamAccountName)) {
                $shape = $page.Drop($userShape, $xPos * 2, 10 - $yPos)
                $shape.Text = "$($_.UserName)`n($($_.SamAccountName))"
                $shapes[$_.SamAccountName] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }
        
        # Render Groups
        $yPos = 1
        $xPos += 2
        $Data | Sort-Object Group -Unique | ForEach-Object {
            if ($_.Group -ne "No Groups" -and -not $shapes.ContainsKey($_.Group)) {
                $shape = $page.Drop($groupShape, $xPos * 2, 10 - $yPos)
                $shape.Text = $_.Group
                $shapes[$_.Group] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }
        
        # Create connect lines
        foreach ($item in $Data) {
            if ($item.Group -ne "No Groups") {
                $page.Shapes.AddConnector(1, $shapes[$item.SamAccountName], $shapes[$item.Group]) | Out-Null
            }
        }
        
        if (Test-Path $Path) { Remove-Item $Path -Force }
        $doc.SaveAs($Path)
        Write-Host "Visio visualization completed: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "Visio export failed: $_"
    }
    finally {
        if ($visio) {
            $visio.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
        }
    }
}

function Export-ToMindManager {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
    $mm = $null
    try {
        $mm = New-Object -ComObject MindManager.Application
        $mm.Visible = $false
        
        $doc = $mm.Documents.Add()
        $root = $doc.CentralTopic
        $root.Text = "AD Benutzer & Gruppen"
        
        # Group by OU for hierarchy representation
        $Data | Group-Object OU | Sort-Object {
            if ($_.Name -match '^\d{2}') { 
                [int]($_.Name -replace '^(\d{2}).*$','$1')
            } else { 
                999 
            }
        } | ForEach-Object {
            $ouTopic = $root.AddSubTopic()
            $ouTopic.Text = $_.Name
            
            $_.Group | Group-Object SamAccountName | ForEach-Object {
                $userTopic = $ouTopic.AddSubTopic()
                $userTopic.Text = "$($_.Group[0].UserName)`n($($_.Name))"
                
                $_.Group | ForEach-Object {
                    if ($_.Group -ne "No Groups") {
                        $groupTopic = $userTopic.AddSubTopic()
                        $groupTopic.Text = $_.Group
                    }
                }
            }
        }
        
        if (Test-Path $Path) { Remove-Item $Path -Force }
        $doc.SaveAs($Path)
        Write-Host "MindManager visualization completed: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "MindManager export failed: $_"
    }
    finally {
        if ($mm) {
            $mm.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mm) | Out-Null
        }
    }
}

# --- Execution ---
try {
    Write-Host "Starting AD user report generation..." -ForegroundColor Green
    
    if ($Force) {
        Close-ExcelProcesses
    }
    
    $result = Get-ADUserGroupData -Pattern $SearchPattern
    if (-not $result) {
        Write-Host "No data to export." -ForegroundColor Yellow
        return
    }
    
    # Sort dataset
    $sortedData = $result.Data | Sort-Object {
        if ($_.SortPrefix -match '^\d{2}$') { 
            [int]$_.SortPrefix 
        } else { 
            999 
        }
    }, UserName, Group | Select-Object OU, UserName, SamAccountName, Group, Comment
    
    Write-Host "Processing $($sortedData.Count) user-group combinations..." -ForegroundColor Green
    
    # Run exports
    if ($OutputPath) { Export-ToCSV -Data $sortedData -Path $OutputPath }
    if ($ExcelPath) { Export-ToExcel -Data $sortedData -GroupColors $result.Colors -Path $ExcelPath }
    if ($HtmlPath) { Export-ToHtml -Data $sortedData -Path $HtmlPath }
    if ($VisioPath) { Export-ToVisio -Data $sortedData -Path $VisioPath }
    if ($MindManagerPath) { Export-ToMindManager -Data $sortedData -Path $MindManagerPath }
    
    Write-Host "Report generation completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "An error occurred during execution: $_"
}
