# AD User Groups Report Script

# Function to force-close Excel
function Close-Excel {
    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
    Write-Verbose "Closed any open Excel processes."
}

# Function to force delete a file
function Remove-FileForce {
    param([string]$path)
    if (Test-Path $path) {
        Remove-Item -Path $path -Force
        Write-Verbose "Forcefully deleted file: $path"
    }
}

# Main script
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

Write-Verbose "Starting AD User Groups Report Script..."

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName Microsoft.Office.Interop.Visio
Add-Type -AssemblyName Mindjet.MindManager.Interop

# Define output paths
$csvPath = "C:\Output\ADUserGroups.csv"
$excelPath = "C:\Output\ADUserGroups.xlsx"
$htmlPath = "C:\Output\ADUserGroups.html"
$visioPath = "C:\Output\ADUserGroups.vsdx"
$mindManagerPath = "C:\Output\ADUserGroups.mmap"

# Ensure output directory exists
New-Item -ItemType Directory -Force -Path (Split-Path $csvPath)

# Close Excel and force delete existing files
Close-Excel
Remove-FileForce $csvPath
Remove-FileForce $excelPath
Remove-FileForce $htmlPath
Remove-FileForce $visioPath
Remove-FileForce $mindManagerPath

Write-Verbose "Retrieving AD users..."
$users = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties SamAccountName, Name, DistinguishedName, MemberOf, Description

$data = @()
foreach ($user in $users) {
    Write-Verbose "Processing user: $($user.SamAccountName)"
    $ou = ($user.DistinguishedName -split ',OU=')[1]
    if ($ou -match '^\d{2,3}$') {
        foreach ($group in $user.MemberOf) {
            $groupName = (Get-ADGroup $group).Name
            $data += [PSCustomObject]@{
                OU = $ou
                User = $user.Name
                Group = $groupName
                Comment = $user.Description
            }
        }
    }
}

$sortedData = $data | Sort-Object {[int]$_.OU}, User, Group

Write-Verbose "Exporting to CSV..."
$sortedData | Export-Csv -Path $csvPath -NoTypeInformation

Write-Verbose "Creating Excel file..."
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Add headers
$headers = @("OU", "User", "Group", "Comment")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
}

# Add data
for ($i = 0; $i -lt $sortedData.Count; $i++) {
    $worksheet.Cells.Item($i + 2, 1) = $sortedData[$i].OU
    $worksheet.Cells.Item($i + 2, 2) = $sortedData[$i].User
    $worksheet.Cells.Item($i + 2, 3) = $sortedData[$i].Group
    $worksheet.Cells.Item($i + 2, 4) = $sortedData[$i].Comment
    
    # Color group cell
    $worksheet.Cells.Item($i + 2, 3).Interior.ColorIndex = ($sortedData[$i].Group.GetHashCode() % 56) + 1
}

$worksheet.UsedRange.EntireColumn.AutoFit()
$workbook.SaveAs($excelPath)
$excel.Quit()

Write-Verbose "Creating HTML visualization..."
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>AD User Groups</title>
    <script src="https://d3js.org/d3.v5.min.js"></script>
    <style>
        .node { cursor: pointer; }
        .node circle { fill: #fff; stroke: steelblue; stroke-width: 1.5px; }
        .node text { font: 10px sans-serif; }
        .link { fill: none; stroke: #ccc; stroke-width: 1.5px; }
    </style>
</head>
<body>
    <div id="tree"></div>
    <script>
        var data = {
            name: "AD Users",
            children: [
                $($sortedData | Group-Object OU | ForEach-Object {
                    "{name: '$($_.Name)', children: [" + ($_.Group | Group-Object User | ForEach-Object {
                        "{name: '$($_.Name)', children: [" + ($_.Group | ForEach-Object {
                            "{name: '$($_.Group)'}"
                        }) -join ',' + "]}"
                    }) -join ',' + "]}"
                })
            ]
        };

        var width = 960, height = 800;
        var tree = d3.tree().size([height, width - 160]);
        var root = d3.hierarchy(data);
        tree(root);

        var svg = d3.select("#tree").append("svg")
            .attr("width", width)
            .attr("height", height)
            .append("g")
            .attr("transform", "translate(80,0)");

        var link = svg.selectAll(".link")
            .data(root.descendants().slice(1))
            .enter().append("path")
            .attr("class", "link")
            .attr("d", d => `M${d.y},${d.x}C${(d.y + d.parent.y) / 2},${d.x} ${(d.y + d.parent.y) / 2},${d.parent.x} ${d.parent.y},${d.parent.x}`);

        var node = svg.selectAll(".node")
            .data(root.descendants())
            .enter().append("g")
            .attr("class", "node")
            .attr("transform", d => `translate(${d.y},${d.x})`);

        node.append("circle")
            .attr("r", 4.5);

        node.append("text")
            .attr("dy", ".31em")
            .attr("x", d => d.children ? -6 : 6)
            .style("text-anchor", d => d.children ? "end" : "start")
            .text(d => d.data.name);
    </script>
</body>
</html>
"@

$html | Out-File -FilePath $htmlPath

Write-Verbose "Creating Visio diagram..."
$visio = New-Object -ComObject Visio.Application
$document = $visio.Documents.Add("")
$page = $document.Pages.Item(1)

$ouShapes = @{}
$userShapes = @{}
$yOffset = 0

foreach ($ou in ($sortedData | Group-Object OU)) {
    $ouShape = $page.DrawRectangle(1, 10 - $yOffset, 2, 9 - $yOffset)
    $ouShape.Text = "OU: $($ou.Name)"
    $ouShapes[$ou.Name] = $ouShape
    
    foreach ($user in ($ou.Group | Group-Object User)) {
        $userShape = $page.DrawRectangle(3, 10 - $yOffset, 4, 9.5 - $yOffset)
        $userShape.Text = $user.Name
        $userShapes[$user.Name] = $userShape
        $page.DrawConnector($ouShape, $userShape)
        
        foreach ($group in $user.Group) {
            $groupShape = $page.DrawRectangle(5, 10 - $yOffset, 6, 9.75 - $yOffset)
            $groupShape.Text = $group.Group
            $page.DrawConnector($userShape, $groupShape)
            $yOffset += 0.5
        }
        $yOffset += 1
    }
    $yOffset += 2
}

$document.SaveAs($visioPath)
$visio.Quit()

Write-Verbose "Creating MindManager map..."
$mm = New-Object -ComObject MindManager.Application
$mm.Visible = $true
$document = $mm.Documents.Add()
$map = $document.CentralTopic

foreach ($ou in ($sortedData | Group-Object OU)) {
    $ouTopic = $map.AddSubTopic($ou.Name)
    foreach ($user in ($ou.Group | Group-Object User)) {
        $userTopic = $ouTopic.AddSubTopic($user.Name)
        foreach ($group in $user.Group) {
            $userTopic.AddSubTopic($group.Group)
        }
    }
}

$document.SaveAs($mindManagerPath)
$mm.Quit()

Write-Verbose "Script completed successfully."
