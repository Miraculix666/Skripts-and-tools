# Module for creating visualizations
function Export-VisualizationsEnhanced {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object[]]$Data,
        [Parameter(Mandatory)]
        [string]$HtmlPath,
        [Parameter(Mandatory)]
        [string]$VisioPath,
        [Parameter(Mandatory)]
        [string]$MindManagerPath
    )

    try {
        # Export HTML visualization
        Export-HtmlVisualization -Data $Data -Path $HtmlPath

        # Export Visio diagram
        Export-VisioVisualization -Data $Data -Path $VisioPath

        # Export MindManager map
        Export-MindManagerVisualization -Data $Data -Path $MindManagerPath
    }
    catch {
        throw "Error creating visualizations: $_"
    }
}

function Export-HtmlVisualization {
    param ($Data, $Path)

    # Process and build a complete force-directed graph with full styling and drag behavior
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>AD User Groups Visualization</title>
    <script src="https://d3js.org/d3.v7.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #1a1a2e; color: #e2e2e2; }
        h1 { text-align: center; color: #00f2fe; margin-bottom: 20px; font-weight: 300; }
        .node circle { stroke: #fff; stroke-width: 2px; shadow: 0 4px 8px rgba(0,0,0,0.2); }
        .node text { font: 12px 'Segoe UI', sans-serif; fill: #e0e0e0; pointer-events: none; text-shadow: 0 1px 3px rgba(0,0,0,0.8); }
        .link { fill: none; stroke: #4f5b66; stroke-opacity: 0.6; stroke-width: 1.5px; }
        #visualization { width: 100%; height: 800px; border: 1px solid #2e3440; border-radius: 12px; background-color: #161623; box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
    </style>
</head>
<body>
    <h1>AD User Groups Visualization</h1>
    <div id="visualization"></div>
    <script>
        const data = {
            nodes: [],
            links: []
        };
        
        // Process data for visualization
        const users = $($Data | ConvertTo-Json);
        const processedUsers = new Set();
        const processedGroups = new Set();
        
        users.forEach(user => {
            const userNameVal = user.Benutzer || user.UserName;
            const groupVal = user.Gruppe || user.Group;
            
            if (!processedUsers.has(user.SamAccountName)) {
                data.nodes.push({
                    id: user.SamAccountName,
                    type: 'user',
                    name: userNameVal,
                    ou: user.OU
                });
                processedUsers.add(user.SamAccountName);
            }
            
            if (!processedGroups.has(groupVal)) {
                data.nodes.push({
                    id: groupVal,
                    type: 'group',
                    name: groupVal
                });
                processedGroups.add(groupVal);
            }
            
            data.links.push({
                source: user.SamAccountName,
                target: groupVal
            });
        });
        
        // Create force-directed graph
        const width = document.getElementById("visualization").clientWidth || window.innerWidth - 40;
        const height = 800;
        
        const simulation = d3.forceSimulation(data.nodes)
            .force("link", d3.forceLink(data.links).id(d => d.id).distance(120))
            .force("charge", d3.forceManyBody().strength(-350))
            .force("center", d3.forceCenter(width / 2, height / 2))
            .force("collision", d3.forceCollide().radius(25));
        
        const svg = d3.select("#visualization")
            .append("svg")
            .attr("width", width)
            .attr("height", height);
            
        // Add zoom and pan behavior
        svg.call(d3.zoom().on("zoom", (event) => {
            container.attr("transform", event.transform);
        }));
        
        const container = svg.append("g");
        
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
            .attr("r", d => d.type === 'user' ? 6 : 10)
            .style("fill", d => d.type === 'user' ? "#00b4db" : "#00f2fe")
            .style("stroke", d => d.type === 'user' ? "#0083b0" : "#4facfe");
        
        node.append("text")
            .text(d => d.name)
            .attr("x", 12)
            .attr("y", 4);
        
        simulation.on("tick", () => {
            link
                .attr("x1", d => d.source.x)
                .attr("y1", d => d.source.y)
                .attr("x2", d => d.target.x)
                .attr("y2", d => d.target.y);
            
            node
                .attr("transform", d => `translate(\${d.x},\${d.y})`);
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
    Write-Host "HTML visualization exported: $Path" -ForegroundColor Green
}

function Export-VisioVisualization {
    param ($Data, $Path)

    $visio = $null
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false

        # Check if file exists and open it, otherwise create new
        if (Test-Path $Path) {
            $doc = $visio.Documents.Open($Path)
            $page = $doc.Pages.Item(1)
            # Clear existing shapes but maintain the document
            $page.Shapes | ForEach-Object { $_.Delete() }
        } else {
            $doc = $visio.Documents.Add("")
            $page = $doc.Pages.Item(1)
        }

        # Try to open stencil for shapes supporting both metric/US and VSX/VSS file formats
        $stencil = $null
        $stencilNames = @("ORGCH_M.VSS", "ORGCH_U.VSS", "ORGCH_M.VSX", "ORGCH_U.VSX")
        foreach ($sName in $stencilNames) {
            try {
                $stencil = $visio.Documents.OpenEx($sName, 4)
                if ($stencil) { break }
            } catch {}
        }

        if (-not $stencil) {
            throw "Could not open Org Chart stencil (ORGCH_M.VSS or similar). Please ensure Visio is installed."
        }

        # Dynamic Shape Lookup (supports English and German stencil shape names)
        $userShape = $null
        $userShapeNames = @("Process", "Prozess", "Task", "Aufgabe")
        foreach ($uName in $userShapeNames) {
            try {
                $userShape = $stencil.Masters.Item($uName)
                if ($userShape) { break }
            } catch {}
        }

        $groupShape = $null
        $groupShapeNames = @("Decision", "Entscheidung", "Process", "Prozess")
        foreach ($gName in $groupShapeNames) {
            try {
                $groupShape = $stencil.Masters.Item($gName)
                if ($groupShape) { break }
            } catch {}
        }

        # Fallbacks to first available masters if specific names do not match localized Visio
        if (-not $userShape -and $stencil.Masters.Count -gt 0) { $userShape = $stencil.Masters.Item(1) }
        if (-not $groupShape -and $stencil.Masters.Count -gt 1) { $groupShape = $stencil.Masters.Item(2) }

        # Store shapes with their identifiers
        $shapes = @{}
        $yPos = 1
        $xPos = 1

        # Create user shapes with data links
        $Data | Sort-Object OU, UserName -Unique | ForEach-Object {
            $uNameVal = if ($_.Benutzer) { $_.Benutzer } else { $_.UserName }
            if (-not $shapes.ContainsKey($_.SamAccountName)) {
                $shape = $page.Drop($userShape, $xPos * 2.5, 12 - $yPos * 1.2)
                $shape.Text = "$($uNameVal)`n($($_.SamAccountName))"

                # Add shape custom properties safely (wrapped to support localized/strict stencils)
                try {
                    $shape.AddSection(243) | Out-Null # visSectionProp
                    $rowIdx1 = $shape.AddRow(243, 0, 0)
                    $shape.CellsSRC(243, $rowIdx1, 0).FormulaU = """$($uNameVal)"""
                    $shape.CellsSRC(243, $rowIdx1, 2).FormulaU = """UserName""" # Label

                    $rowIdx2 = $shape.AddRow(243, 1, 0)
                    $shape.CellsSRC(243, $rowIdx2, 0).FormulaU = """$($_.SamAccountName)"""
                    $shape.CellsSRC(243, $rowIdx2, 2).FormulaU = """SamAccountName""" # Label
                } catch {
                    Write-Verbose "Failed to attach custom shape properties to User: $($_.SamAccountName)"
                }

                $shapes[$_.SamAccountName] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }

        # Create group shapes with data links
        $yPos = 1
        $xPos += 2.5
        $Data | Sort-Object Group -Unique | ForEach-Object {
            $groupVal = if ($_.Gruppe) { $_.Gruppe } else { $_.Group }
            if (-not $shapes.ContainsKey($groupVal)) {
                $shape = $page.Drop($groupShape, $xPos * 2.5, 12 - $yPos * 1.2)
                $shape.Text = $groupVal

                # Add shape custom properties safely
                try {
                    $shape.AddSection(243) | Out-Null # visSectionProp
                    $rowIdx1 = $shape.AddRow(243, 0, 0)
                    $shape.CellsSRC(243, $rowIdx1, 0).FormulaU = """$($groupVal)"""
                    $shape.CellsSRC(243, $rowIdx1, 2).FormulaU = """GroupName""" # Label
                } catch {
                    Write-Verbose "Failed to attach custom shape properties to Group: $groupVal"
                }

                $shapes[$groupVal] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }

        # Add connections safely
        $Data | ForEach-Object {
            $groupVal = if ($_.Gruppe) { $_.Gruppe } else { $_.Group }
            try {
                $page.Shapes.AddConnector(1, $shapes[$_.SamAccountName], $shapes[$groupVal]) | Out-Null
            } catch {
                Write-Verbose "Failed to connect user $($_.SamAccountName) to group $groupVal"
            }
        }

        # Save and close
        $doc.SaveAs($Path)
        $doc.Close()
        Write-Host "Visio visualization exported: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "Visio export failed: $_"
    }
    finally {
        if ($visio) {
            try { $visio.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

function Export-MindManagerVisualization {
    param ($Data, $Path)

    $mm = $null
    try {
        $mm = New-Object -ComObject MindManager.Application
        $mm.Visible = $false

        # Check if file exists and open it, otherwise create new
        if (Test-Path $Path) {
            $doc = $mm.Documents.Open($Path)
            # Clear existing topics but keep root
            $root = $doc.CentralTopic
            $root.SubTopics | ForEach-Object { $_.Delete() }
        } else {
            $doc = $mm.Documents.Add()
            $root = $doc.CentralTopic
            $root.Text = "AD User Groups"
        }

        # Safe marker configuration to prevent errors in differing API versions
        $userMarker = $null
        $groupMarker = $null
        try {
            $doc.CustomProperties.Add("LastUpdate", (Get-Date).ToString())
            $markers = $doc.Markers
            $userMarker = $markers.Add("User")
            $groupMarker = $markers.Add("Group")
        }
        catch {
            Write-Verbose "MindManager markers not supported or failed to create: $_"
        }

        # Group by OU
        $Data | Group-Object OU | Sort-Object {
            if ($_.Name -match '^\d{2}') {
                [int]($_.Name -replace '^(\d{2}).*$','$1')
            } else {
                999
            }
        } | ForEach-Object {
            $ouTopic = $root.AddSubTopic()
            $ouTopic.Text = $_.Name

            # Store OU data safely
            try { $ouTopic.Properties.Add("OU_ID", $_.Name) | Out-Null } catch {}

            # Add users under OU
            $_.Group | Group-Object SamAccountName | ForEach-Object {
                $uNameVal = if ($_.Group[0].Benutzer) { $_.Group[0].Benutzer } else { $_.Group[0].UserName }
                $userTopic = $ouTopic.AddSubTopic()
                $userTopic.Text = "$($uNameVal)`n($($_.Name))"
                
                if ($userMarker) {
                    try { $userTopic.AddMarker($userMarker) | Out-Null } catch {}
                }

                # Store user data safely
                try {
                    $userTopic.Properties.Add("SamAccountName", $_.Name) | Out-Null
                    $userTopic.Properties.Add("UserName", $uNameVal) | Out-Null
                } catch {}

                # Add groups under user
                $_.Group | ForEach-Object {
                    $groupVal = if ($_.Gruppe) { $_.Gruppe } else { $_.Group }
                    $groupTopic = $userTopic.AddSubTopic()
                    $groupTopic.Text = $groupVal
                    
                    if ($groupMarker) {
                        try { $groupTopic.AddMarker($groupMarker) | Out-Null } catch {}
                    }

                    # Store group data safely
                    try { $groupTopic.Properties.Add("GroupName", $groupVal) | Out-Null } catch {}
                }
            }
        }

        # Save map and close
        $doc.SaveAs($Path)
        $doc.Close()
        Write-Host "MindManager visualization exported: $Path" -ForegroundColor Green
    }
    catch {
        Write-Warning "MindManager export failed: $_"
    }
    finally {
        if ($mm) {
            try { $mm.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mm) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

Export-ModuleMember -Function Export-VisualizationsEnhanced
