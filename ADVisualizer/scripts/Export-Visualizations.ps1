function Export-ToVisio {
    param (
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Path
    )
    
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
        
        # Create stencil for shapes
        $stencil = $visio.Documents.OpenEx("ORGCH_M.VSS", 4)
        $userShape = $stencil.Masters.Item("Process")
        $groupShape = $stencil.Masters.Item("Decision")
        
        # Create data recordset for shape data
        $dataRecordset = $doc.DataRecordsets.Add()
        $dataRecordset.Name = "ADUserGroups"
        
        # Add data columns
        $dataRecordset.DataColumns.Add("UserName", "UserName")
        $dataRecordset.DataColumns.Add("SamAccountName", "SamAccountName")
        $dataRecordset.DataColumns.Add("Group", "Group")
        
        # Store shapes with their identifiers
        $shapes = @{}
        $yPos = 1
        $xPos = 1
        
        # Create user shapes with data links
        $Data | Sort-Object OU, UserName -Unique | ForEach-Object {
            if (-not $shapes.ContainsKey($_.SamAccountName)) {
                $shape = $page.Drop($userShape, $xPos * 2, 10 - $yPos)
                $shape.Text = "$($_.UserName)`n($($_.SamAccountName))"
                
                # Add shape data
                $shape.AddRow("User_Data", 0)
                $shape.Cells("User_Data.UserName").Formula = """$($_.UserName)"""
                $shape.Cells("User_Data.SamAccountName").Formula = """$($_.SamAccountName)"""
                
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
        $xPos += 2
        $Data | Sort-Object Group -Unique | ForEach-Object {
            if (-not $shapes.ContainsKey($_.Group)) {
                $shape = $page.Drop($groupShape, $xPos * 2, 10 - $yPos)
                $shape.Text = $_.Group
                
                # Add shape data
                $shape.AddRow("Group_Data", 0)
                $shape.Cells("Group_Data.GroupName").Formula = """$($_.Group)"""
                
                $shapes[$_.Group] = $shape
                $yPos++
                if ($yPos -gt 8) {
                    $yPos = 1
                    $xPos++
                }
            }
        }
        
        # Add connections with data links
        $Data | ForEach-Object {
            $connector = $page.Shapes.AddConnector(1, $shapes[$_.SamAccountName], $shapes[$_.Group])
            $connector.AddRow("Connection_Data", 0)
            $connector.Cells("Connection_Data.FromUser").Formula = """$($_.SamAccountName)"""
            $connector.Cells("Connection_Data.ToGroup").Formula = """$($_.Group)"""
        }
        
        # Save with data links
        $doc.SaveAs($Path)
        $visio.Quit()
        
        Write-Host "Visio visualization exported: $Path" -ForegroundColor Green
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
        
        # Add custom properties for data tracking
        $doc.CustomProperties.Add("LastUpdate", (Get-Date).ToString())
        
        # Create markers for different types
        $markers = $doc.Markers
        $userMarker = $markers.Add("User")
        $groupMarker = $markers.Add("Group")
        
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
            
            # Store OU data for updates
            $ouTopic.Properties.Add("OU_ID", $_.Name)
            
            # Add users under OU
            $_.Group | Group-Object SamAccountName | ForEach-Object {
                $userTopic = $ouTopic.AddSubTopic()
                $userTopic.Text = "$($_.Group[0].UserName)`n($($_.Name))"
                $userTopic.AddMarker($userMarker)
                
                # Store user data for updates
                $userTopic.Properties.Add("SamAccountName", $_.Name)
                $userTopic.Properties.Add("UserName", $_.Group[0].UserName)
                
                # Add groups under user
                $_.Group | ForEach-Object {
                    $groupTopic = $userTopic.AddSubTopic()
                    $groupTopic.Text = $_.Group
                    $groupTopic.AddMarker($groupMarker)
                    
                    # Store group data for updates
                    $groupTopic.Properties.Add("GroupName", $_.Group)
                }
            }
        }
        
        # Save with update capability
        $doc.SaveAs($Path)
        $mm.Quit()
        
        Write-Host "MindManager visualization exported: $Path" -ForegroundColor Green
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