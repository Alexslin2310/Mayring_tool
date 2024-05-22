Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Global variables to store categories and subcategories
$globalCategories = @()
$globalSubCategories = @()

# Function to load and analyze JSON files
function Load-And-Analyze-JSON {
    param (
        [string]$folderPath,
        [ref]$progressBar
    )

    $categories = @{}

    $jsonFiles = Get-ChildItem -Path $folderPath -Filter *.json
    $totalFiles = $jsonFiles.Count
    $progressBar.Value.Maximum = $totalFiles
    $progressBar.Value.Value = 0

    foreach ($file in $jsonFiles) {
        $jsonContent = Get-Content -Path $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($item in $jsonContent) {
            # Convert the item to a PowerShell object and add the JsonFile property
            $item = [PSCustomObject]@{
                ID               = [int]$item.ID
                Nummer           = [int]$item.Nummer
                Textausschnitt   = $item.Textausschnitt
                Paraphrase       = if ($item.Paraphrase) { $item.Paraphrase } else { "" }
                Generalisierung  = if ($item.Generalisierung) { $item.Generalisierung } else { "" }
                Kategorie        = if ($item.Kategorie) { $item.Kategorie } else { "" }
                Subkategorie     = if ($item.Subkategorie) { $item.Subkategorie } else { "" }
                JsonFile         = $file.FullName
            }

            $category = $item.Kategorie
            $subCategory = $item.Subkategorie

            if (-not $categories.ContainsKey($category)) {
                $categories[$category] = @{}
                if ($category -ne $null -and $category -ne "") {
                    $globalCategories += $category
                }
            }

            if (-not $categories[$category].ContainsKey($subCategory)) {
                $categories[$category][$subCategory] = @()
                if ($subCategory -ne $null -and $subCategory -ne "") {
                    $globalSubCategories += $subCategory
                }
            }

            $categories[$category][$subCategory] += $item
        }

        $progressBar.Value.Dispatcher.Invoke([action]{ $progressBar.Value.Value += 1 }, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    return $categories
}

# Function to load categories and subcategories from JSON file
function Load-CategoryData {
    param (
        [string]$filePath
    )

    $categoryData = Get-Content -Path $filePath -Raw -Encoding UTF8 | ConvertFrom-Json
    return @{
        "Kategorien" = $categoryData.Kategorien
        "Subkategorien" = $categoryData.Subkategorien
    }
}

# Function to populate TreeView with analysis results
function Populate-TreeView {
    param (
        [System.Windows.Controls.TreeView]$treeView,
        [hashtable]$categories
    )

    $treeView.Items.Clear()
    
    $sortedCategories = $categories.GetEnumerator() | Sort-Object -Property {
        $_.Value.Values | ForEach-Object { $_.Count } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    } -Descending

    foreach ($categoryEntry in $sortedCategories) {
        $category = $categoryEntry.Key
        $categoryItems = @()
        foreach ($subItems in $categories[$category].Values) {
            $categoryItems += $subItems
        }

        # Sort category items by ID and Nummer
        $categoryItems = $categoryItems | Sort-Object -Property ID, Nummer

        $categoryNode = New-Object System.Windows.Controls.TreeViewItem
        $categoryNode.Header = "$category (Count: $($categoryItems.Count))"
        $categoryNode.Tag = @{"Type"="Category"; "Items"=$categoryItems}
        $treeView.Items.Add($categoryNode)

        $sortedSubCategories = $categories[$category].GetEnumerator() | Sort-Object -Property { $_.Value.Count } -Descending

        foreach ($subCategoryEntry in $sortedSubCategories) {
            $subCategory = $subCategoryEntry.Key
            $subCategoryItems = $categories[$category][$subCategory]

            # Sort subcategory items by ID and Nummer
            $subCategoryItems = $subCategoryItems | Sort-Object -Property ID, Nummer

            $subCategoryNode = New-Object System.Windows.Controls.TreeViewItem
            $subCategoryNode.Header = "$subCategory (Count: $($subCategoryItems.Count))"
            $subCategoryNode.Tag = @{"Type"="SubCategory"; "Items"=$subCategoryItems}
            $categoryNode.Items.Add($subCategoryNode)
        }
    }

    foreach ($node in $treeView.Items) {
        $node.ExpandSubtree()
    }
}

# Function to save the edited items back to the JSON files and refresh the TreeView
function Save-Items {
    param (
        [array]$items,
        [hashtable]$categories,
        [System.Windows.Controls.TreeView]$treeView
    )

    $groupedByFile = $items | Group-Object -Property JsonFile

    foreach ($group in $groupedByFile) {
        $filePath = $group.Name
        $jsonContent = Get-Content -Path $filePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $fileItems = $group.Group

        foreach ($item in $fileItems) {
            $jsonItem = $jsonContent | Where-Object { $_.ID -eq $item.ID -and $_.Nummer -eq $item.Nummer }
            if ($jsonItem -ne $null) {
                $jsonItem.Textausschnitt = $item.Textausschnitt
                $jsonItem.Paraphrase = $item.Paraphrase
                $jsonItem.Generalisierung = $item.Generalisierung
                $jsonItem.Kategorie = $item.Kategorie
                $jsonItem.Subkategorie = $item.Subkategorie
            }
        }

        $jsonContent | ConvertTo-Json -Depth 3 | Set-Content -Path $filePath -Encoding UTF8
    }

    # Reload categories and refresh TreeView only if treeView is not null
    if ($treeView -ne $null) {
        $folderPath = Split-Path -Parent $groupedByFile[0].Name
        $progressBar = New-Object System.Windows.Controls.ProgressBar
        $categories = Load-And-Analyze-JSON -folderPath $folderPath -progressBar ([ref]$progressBar)
        Populate-TreeView -treeView $treeView -categories $categories
    }
}

# Function to export the current items to a CSV file
function Export-ItemsToCSV {
    param (
        [array]$items
    )

    # Remove JsonFile property from items before exporting
    $itemsToExport = $items | Select-Object -Property ID, Nummer, Textausschnitt, Paraphrase, Generalisierung, Kategorie, Subkategorie

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save CSV File"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $itemsToExport | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        [System.Windows.MessageBox]::Show("CSV file saved successfully.", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
}

# Function to export the current items to a JSON file
function Export-ItemsToJSON {
    param (
        [array]$items
    )

    # Remove JsonFile property from items before exporting
    $itemsToExport = $items | Select-Object -Property ID, Nummer, Textausschnitt, Paraphrase, Generalisierung, Kategorie, Subkategorie

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "JSON files (*.json)|*.json"
    $saveFileDialog.Title = "Save JSON File"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $itemsToExport | ConvertTo-Json -Depth 3 | Set-Content -Path $filePath -Encoding UTF8
        [System.Windows.MessageBox]::Show("JSON file saved successfully.", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
}

# Function to create the details window with editable textboxes and combo boxes
function Show-DetailsWindow {
    param (
        [string]$header,
        [array]$items,
        [hashtable]$categories,
        [System.Windows.Controls.TreeView]$treeView,
        [hashtable]$categoryData
    )

    if ($items.Count -gt 1) {
        $sortedItems = $items | Sort-Object -Property ID, Nummer
    } else {
        $sortedItems = $items
    }

    $detailsWindow = New-Object System.Windows.Window
    $detailsWindow.Title = "Details"
    $detailsWindow.Height = 600
    $detailsWindow.Width = 800
    $detailsWindow.WindowStartupLocation = 'CenterScreen'

    $grid = New-Object System.Windows.Controls.Grid

    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = [System.Windows.GridLength]::Auto}))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*"}))
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = [System.Windows.GridLength]::Auto}))

    $headerTextBlock = New-Object System.Windows.Controls.TextBlock
    $headerTextBlock.Text = $header
    $headerTextBlock.FontWeight = 'Bold'
    $headerTextBlock.FontSize = 16
    $headerTextBlock.Margin = 10

    $dataGrid = New-Object System.Windows.Controls.DataGrid
    $dataGrid.Margin = 10
    $dataGrid.AutoGenerateColumns = $false
    $dataGrid.CanUserAddRows = $false
    $dataGrid.IsReadOnly = $false

    # Alternating row colors
    $dataGrid.AlternatingRowBackground = [System.Windows.Media.Brushes]::LightBlue
    $dataGrid.RowBackground = [System.Windows.Media.Brushes]::AliceBlue

    $idColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $idColumn.Header = "ID"
    $idColumn.Binding = New-Object System.Windows.Data.Binding("ID")
    $idColumn.Width = 50
    $dataGrid.Columns.Add($idColumn)

    $nummerColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $nummerColumn.Header = "Nummer"
    $nummerColumn.Binding = New-Object System.Windows.Data.Binding("Nummer")
    $nummerColumn.Width = 100
    $dataGrid.Columns.Add($nummerColumn)

    $textausschnittColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $textausschnittColumn.Header = "Textausschnitt"
    $textausschnittColumn.Binding = New-Object System.Windows.Data.Binding("Textausschnitt")
    $textausschnittColumn.Width = 300
    $textausschnittColumn.ElementStyle = [Windows.Markup.XamlReader]::Parse(
        "<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='TextBlock'>
            <Setter Property='TextWrapping' Value='Wrap' />
            <Setter Property='Margin' Value='5' />
        </Style>"
    )
    $dataGrid.Columns.Add($textausschnittColumn)

    $paraphraseColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $paraphraseColumn.Header = "Paraphrase"
    $paraphraseColumn.Binding = New-Object System.Windows.Data.Binding("Paraphrase")
    $paraphraseColumn.Width = 300
    $paraphraseColumn.ElementStyle = [Windows.Markup.XamlReader]::Parse(
        "<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='TextBlock'>
            <Setter Property='TextWrapping' Value='Wrap' />
            <Setter Property='Margin' Value='5' />
        </Style>"
    )
    $dataGrid.Columns.Add($paraphraseColumn)

    $generalisierungColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $generalisierungColumn.Header = "Generalisierung"
    $generalisierungColumn.Binding = New-Object System.Windows.Data.Binding("Generalisierung")
    $generalisierungColumn.Width = 300
    $generalisierungColumn.ElementStyle = [Windows.Markup.XamlReader]::Parse(
        "<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='TextBlock'>
            <Setter Property='TextWrapping' Value='Wrap' />
            <Setter Property='Margin' Value='5' />
        </Style>"
    )
    $dataGrid.Columns.Add($generalisierungColumn)

    $kategorieColumn = New-Object System.Windows.Controls.DataGridComboBoxColumn
    $kategorieColumn.Header = "Kategorie"
    $kategorieColumn.SelectedItemBinding = New-Object System.Windows.Data.Binding("Kategorie")
    $kategorieColumn.ItemsSource = $categoryData.Kategorien
    $kategorieColumn.Width = 200
    $dataGrid.Columns.Add($kategorieColumn)

    $subkategorieColumn = New-Object System.Windows.Controls.DataGridComboBoxColumn
    $subkategorieColumn.Header = "Subkategorie"
    $subkategorieColumn.SelectedItemBinding = New-Object System.Windows.Data.Binding("Subkategorie")
    $subkategorieColumn.ItemsSource = $categoryData.Subkategorien
    $subkategorieColumn.Width = 200
    $dataGrid.Columns.Add($subkategorieColumn)

    $dataGrid.ItemsSource = $sortedItems

    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Margin = "10"

    $buttonStyle = [Windows.Markup.XamlReader]::Parse(
        "<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='Button'>
            <Setter Property='Background' Value='Gray' />
            <Setter Property='Foreground' Value='White' />
            <Setter Property='FontWeight' Value='Bold' />
            <Setter Property='BorderBrush' Value='DarkGray' />
            <Setter Property='BorderThickness' Value='2' />
            <Setter Property='Padding' Value='5,2' />
            <Setter Property='Margin' Value='0,0,10,0' />
            <Setter Property='Cursor' Value='Hand' />
        </Style>"
    )

    $saveButton = New-Object System.Windows.Controls.Button
    $saveButton.Content = "Save"
    $saveButton.Style = $buttonStyle
    $saveButton.Add_Click({
        Save-Items -items $sortedItems -categories $categories -treeView $treeView
        [System.Windows.MessageBox]::Show("Changes saved successfully.", "Save Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    $exportCSVButton = New-Object System.Windows.Controls.Button
    $exportCSVButton.Content = "Export to CSV"
    $exportCSVButton.Style = $buttonStyle
    $exportCSVButton.Add_Click({
        Export-ItemsToCSV -items $sortedItems
    })

    $exportJSONButton = New-Object System.Windows.Controls.Button
    $exportJSONButton.Content = "Export to JSON"
    $exportJSONButton.Style = $buttonStyle
    $exportJSONButton.Add_Click({
        Export-ItemsToJSON -items $sortedItems
    })

    $buttonPanel.Children.Add($saveButton)
    $buttonPanel.Children.Add($exportCSVButton)
    $buttonPanel.Children.Add($exportJSONButton)

    $grid.Children.Add($headerTextBlock)
    [System.Windows.Controls.Grid]::SetRow($headerTextBlock, 0)
    $grid.Children.Add($dataGrid)
    [System.Windows.Controls.Grid]::SetRow($dataGrid, 1)
    $grid.Children.Add($buttonPanel)
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 2)

    $detailsWindow.Content = $grid

    # Add event handler to update TreeView when details window is closed
    $detailsWindow.add_Closed({
        $folderPath = Split-Path -Parent $items[0].JsonFile
        $progressBar = New-Object System.Windows.Controls.ProgressBar
        $categories = Load-And-Analyze-JSON -folderPath $folderPath -progressBar ([ref]$progressBar)
        Populate-TreeView -treeView $treeView -categories $categories
    })

    $detailsWindow.ShowDialog()
}

# Function to show details of a specific JSON file
function Show-JsonDetails {
    param (
        [string]$filePath,
        [hashtable]$categories,
        [System.Windows.Controls.TreeView]$treeView,
        [hashtable]$categoryData
    )

    $jsonContent = Get-Content -Path $filePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $items = @()

    foreach ($item in $jsonContent) {
        $items += [PSCustomObject]@{
            ID               = [int]$item.ID
            Nummer           = [int]$item.Nummer
            Textausschnitt   = $item.Textausschnitt
            Paraphrase       = $item.Paraphrase
            Generalisierung  = $item.Generalisierung
            Kategorie        = $item.Kategorie
            Subkategorie     = $item.Subkategorie
            JsonFile         = $filePath
        }
    }

    Show-DetailsWindow -header "Details for $filePath" -items $items -categories $categories -treeView $treeView -categoryData $categoryData
}

# Function to create the enhanced GUI using WPF
function Create-GUI {
    $form = New-Object System.Windows.Window
    $form.Title = "JSON Frequency Analysis"
    $form.Height = 600
    $form.Width = 800
    $form.WindowStartupLocation = 'CenterScreen'

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = 10

    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row1)
    
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = "*"
    $grid.RowDefinitions.Add($row2)
    
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row3)

    $groupBox = New-Object System.Windows.Controls.GroupBox
    $groupBox.Header = "Select Folder and Analyze"
    $groupBox.Padding = 10
    $groupBox.Margin = 10

    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Orientation = 'Horizontal'

    $selectFolderButton = New-Object System.Windows.Controls.Button
    $selectFolderButton.Content = "Select Folder"
    $selectFolderButton.Width = 120
    $selectFolderButton.Margin = '0,0,10,0'

    $progressBar = New-Object System.Windows.Controls.ProgressBar
    $progressBar.Height = 30
    $progressBar.Width = 400

    $selectJsonButton = New-Object System.Windows.Controls.Button
    $selectJsonButton.Content = "Select JSON"
    $selectJsonButton.Width = 120
    $selectJsonButton.Margin = '0,0,10,0'
    $selectJsonButton.IsEnabled = $false

    $stackPanel.Children.Add($selectFolderButton)
    $stackPanel.Children.Add($progressBar)
    $stackPanel.Children.Add($selectJsonButton)
    $groupBox.Content = $stackPanel
    $grid.Children.Add($groupBox)
    [System.Windows.Controls.Grid]::SetRow($groupBox, 0)

    $treeView = New-Object System.Windows.Controls.TreeView
    $treeView.Margin = 10
    [System.Windows.Controls.ScrollViewer]::SetHorizontalScrollBarVisibility($treeView, 'Auto')
    [System.Windows.Controls.ScrollViewer]::SetVerticalScrollBarVisibility($treeView, 'Auto')

    $grid.Children.Add($treeView)
    [System.Windows.Controls.Grid]::SetRow($treeView, 1)

    $statusLabel = New-Object System.Windows.Controls.Label
    $statusLabel.Content = "Status: Waiting for folder selection..."
    $statusLabel.Margin = 10
    $grid.Children.Add($statusLabel)
    [System.Windows.Controls.Grid]::SetRow($statusLabel, 2)

    $form.Content = $grid

    $treeView.Add_MouseDoubleClick({
        $selectedItem = $treeView.SelectedItem
        if ($selectedItem -and $selectedItem.Tag) {
            $header = $selectedItem.Header.ToString()
            $items = $selectedItem.Tag["Items"]
            $categoryData = Load-CategoryData -filePath "$PSScriptRoot\Config\Config.Kategorie.json"
            Show-DetailsWindow -header $header -items $items -categories $categories -treeView $treeView -categoryData $categoryData
        }
    })

    $selectFolderButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $folderPath = $folderBrowser.SelectedPath
            $statusLabel.Content = "Status: Analyzing JSON files..."
            $progressBar.Value = 0

            $globalCategories.Clear()
            $globalSubCategories.Clear()

            $categories = Load-And-Analyze-JSON -folderPath $folderPath -progressBar ([ref]$progressBar)
            Populate-TreeView -treeView $treeView -categories $categories

            $statusLabel.Content = "Status: Analysis complete."
            $selectJsonButton.IsEnabled = $true
        }
    })

    $selectJsonButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "JSON files (*.json)|*.json"
        $openFileDialog.Title = "Select a JSON File"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePath = $openFileDialog.FileName
            $categoryData = Load-CategoryData -filePath "$PSScriptRoot\Config\Config.Kategorie.json"
            Show-JsonDetails -filePath $filePath -categories $categories -treeView $treeView -categoryData $categoryData
        }
    })

    $form.ShowDialog()
}

# Run the enhanced GUI
Create-GUI
