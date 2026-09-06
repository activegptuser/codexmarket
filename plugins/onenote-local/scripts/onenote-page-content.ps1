function New-OneNoteNamespaceManager {
    param([xml]$Document)

    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
    $namespaceManager.AddNamespace("one", $Document.DocumentElement.NamespaceURI)
    return ,$namespaceManager
}

function ConvertFrom-OneNoteHtmlText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $withoutTags = [regex]::Replace($Text, "<[^>]+>", "")
    $decoded = [System.Net.WebUtility]::HtmlDecode($withoutTags)
    return ([regex]::Replace($decoded, "\s+", " ")).Trim()
}

function Get-OneNoteNodePlainText {
    param([System.Xml.XmlNode]$Node)

    if ($null -eq $Node) {
        return ""
    }

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($textNode in $Node.SelectNodes(".//*[local-name()='T']")) {
        $text = ConvertFrom-OneNoteHtmlText -Text $textNode.InnerText
        if ($text.Length -gt 0) {
            $parts.Add($text)
        }
    }

    return ($parts -join [Environment]::NewLine)
}

function Get-OneNoteDirectPlainText {
    param(
        [System.Xml.XmlNode]$Node,
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($textNode in $Node.SelectNodes("./one:T", $NamespaceManager)) {
        $text = ConvertFrom-OneNoteHtmlText -Text $textNode.InnerText
        if ($text.Length -gt 0) {
            $parts.Add($text)
        }
    }

    return ($parts -join [Environment]::NewLine)
}

function Get-KanbanWorkHeadingText {
    $codePoints = @(0xC791, 0xC5C5, 0x28, 0xC138, 0xBD80, 0x20, 0xB0B4, 0xC6A9, 0x29)
    return -join @($codePoints | ForEach-Object { [char]$_ })
}

function ConvertFrom-OneNoteCodePoints {
    param([int[]]$CodePoints)

    return -join @($CodePoints | ForEach-Object { [char]$_ })
}

function Get-KanbanSectionHeadingText {
    param([string]$Section)

    switch ($Section) {
        "work" {
            return Get-KanbanWorkHeadingText
        }
        "completion" {
            return ConvertFrom-OneNoteCodePoints @(0xC644, 0xB8CC, 0xC870, 0xAC74)
        }
        "prerequisites" {
            return ConvertFrom-OneNoteCodePoints @(0xC120, 0xD589, 0xC5C5, 0xBB34)
        }
        "risks" {
            return (ConvertFrom-OneNoteCodePoints @(0xB9AC, 0xC2A4, 0xD06C)) + " / " +
                (ConvertFrom-OneNoteCodePoints @(0xC774, 0xC288))
        }
        default {
            throw "Unsupported Kanban section: $Section"
        }
    }
}

function Get-KanbanFieldLabel {
    param([string]$FieldName)

    switch ($FieldName) {
        "assignee" { return ConvertFrom-OneNoteCodePoints @(0xB2F4, 0xB2F9) }
        "workType" { return ConvertFrom-OneNoteCodePoints @(0xC5C5, 0xBB34, 0xC720, 0xD615) }
        "priority" { return ConvertFrom-OneNoteCodePoints @(0xC6B0, 0xC120, 0xC21C, 0xC704) }
        "effort" { return ConvertFrom-OneNoteCodePoints @(0xACF5, 0xC218) }
        "createdDate" { return ConvertFrom-OneNoteCodePoints @(0xC0DD, 0xC131, 0xC77C) }
        "dueDate" { return ConvertFrom-OneNoteCodePoints @(0xB9C8, 0xAC10, 0xC77C) }
        default { throw "Unsupported Kanban field: $FieldName" }
    }
}

function Remove-KanbanHeadingPrefix {
    param([string]$Text)

    $blackSquare = [string][char]0x25A0
    $normalized = $Text.Trim()
    while ($normalized.StartsWith($blackSquare)) {
        $normalized = $normalized.Substring(1).TrimStart()
    }

    return $normalized
}

function Find-UniqueKanbanSectionHeading {
    param(
        [xml]$Document,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$Section
    )

    $targetText = Get-KanbanSectionHeadingText -Section $Section
    $matches = New-Object System.Collections.Generic.List[System.Xml.XmlNode]

    foreach ($oeNode in $Document.SelectNodes("//one:OE[one:T]", $NamespaceManager)) {
        $headingText = Get-OneNoteDirectPlainText -Node $oeNode -NamespaceManager $NamespaceManager
        if ((Remove-KanbanHeadingPrefix -Text $headingText) -eq $targetText) {
            $matches.Add($oeNode)
        }
    }

    if ($matches.Count -ne 1) {
        throw "Expected exactly one Kanban '$Section' heading; found $($matches.Count)."
    }

    return $matches[0]
}

function Find-UniqueKanbanWorkHeading {
    param(
        [xml]$Document,
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    return Find-UniqueKanbanSectionHeading `
        -Document $Document `
        -NamespaceManager $NamespaceManager `
        -Section "work"
}

function Test-OneNoteEntryMarker {
    param(
        [System.Xml.XmlNode]$RootNode,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$EntryId
    )

    foreach ($metaNode in $RootNode.SelectNodes(".//one:Meta", $NamespaceManager)) {
        if (
            $metaNode.GetAttribute("name") -eq "onenote-local.entry-id" -and
            $metaNode.GetAttribute("content") -eq $EntryId
        ) {
            return $true
        }
    }

    return $false
}

function Add-OneNoteEntryMarker {
    param(
        [xml]$Document,
        [System.Xml.XmlNode]$EntryNode,
        [string]$NamespaceUri,
        [string]$EntryId
    )

    $metaNode = $Document.CreateElement("one", "Meta", $NamespaceUri)
    $metaNode.SetAttribute("name", "onenote-local.entry-id")
    $metaNode.SetAttribute("content", $EntryId)

    if ($null -eq $EntryNode.FirstChild) {
        [void]$EntryNode.AppendChild($metaNode)
    }
    else {
        [void]$EntryNode.InsertBefore($metaNode, $EntryNode.FirstChild)
    }
}

function Set-OneNoteEntryText {
    param(
        [xml]$Document,
        [System.Xml.XmlNode]$EntryNode,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$NamespaceUri,
        [string]$Text
    )

    $newTextNode = $Document.CreateElement("one", "T", $NamespaceUri)
    [void]$newTextNode.AppendChild($Document.CreateCDataSection($Text))
    $existingTextNode = $EntryNode.SelectSingleNode("./one:T", $NamespaceManager)

    if ($null -eq $existingTextNode) {
        [void]$EntryNode.AppendChild($newTextNode)
    }
    else {
        [void]$EntryNode.ReplaceChild($newTextNode, $existingTextNode)
    }
}

function New-OneNoteWorklogEntryNode {
    param(
        [xml]$Document,
        [System.Xml.XmlNode]$TemplateNode,
        [string]$NamespaceUri,
        [string]$EntryId,
        [string]$Text
    )

    $entryNode = $Document.CreateElement("one", "OE", $NamespaceUri)
    foreach ($attributeName in @("alignment", "quickStyleIndex", "style")) {
        if ($null -ne $TemplateNode -and $null -ne $TemplateNode.Attributes[$attributeName]) {
            $entryNode.SetAttribute($attributeName, $TemplateNode.Attributes[$attributeName].Value)
        }
    }

    Add-OneNoteEntryMarker `
        -Document $Document `
        -EntryNode $entryNode `
        -NamespaceUri $NamespaceUri `
        -EntryId $EntryId

    if ($null -ne $TemplateNode) {
        $templateListNode = $TemplateNode.SelectSingleNode("./*[local-name()='List']")
        if ($null -ne $templateListNode) {
            [void]$entryNode.AppendChild($Document.ImportNode($templateListNode, $true))
        }
    }

    if ($null -eq $entryNode.SelectSingleNode("./*[local-name()='List']")) {
        $listNode = $Document.CreateElement("one", "List", $NamespaceUri)
        $bulletNode = $Document.CreateElement("one", "Bullet", $NamespaceUri)
        $bulletNode.SetAttribute("bullet", "25")
        $bulletNode.SetAttribute("fontSize", "11.0")
        [void]$listNode.AppendChild($bulletNode)
        [void]$entryNode.AppendChild($listNode)
    }

    $textNode = $Document.CreateElement("one", "T", $NamespaceUri)
    [void]$textNode.AppendChild($Document.CreateCDataSection($Text))
    [void]$entryNode.AppendChild($textNode)
    return $entryNode
}

function Get-OneNoteOutlineAncestor {
    param([System.Xml.XmlNode]$Node)

    $current = $Node
    while ($null -ne $current -and $current.LocalName -ne "Outline") {
        $current = $current.ParentNode
    }

    if ($null -eq $current) {
        throw "The Kanban work details heading is not inside an Outline."
    }

    return $current
}

function Find-UniqueKanbanFieldNode {
    param(
        [xml]$Document,
        [System.Xml.XmlNamespaceManager]$NamespaceManager,
        [string]$FieldName
    )

    $label = Get-KanbanFieldLabel -FieldName $FieldName
    $pattern = "^" + [regex]::Escape($label) + "\s*:\s*"
    $fieldMatches = New-Object System.Collections.Generic.List[System.Xml.XmlNode]

    foreach ($oeNode in $Document.SelectNodes("//one:OE[one:T]", $NamespaceManager)) {
        $directText = Get-OneNoteDirectPlainText -Node $oeNode -NamespaceManager $NamespaceManager
        if ($directText -match $pattern) {
            $fieldMatches.Add($oeNode)
        }
    }

    if ($fieldMatches.Count -ne 1) {
        throw "Expected exactly one Kanban '$FieldName' field; found $($fieldMatches.Count)."
    }

    return $fieldMatches[0]
}

function Update-KanbanFieldsInDocument {
    param(
        [xml]$Document,
        [object]$Fields
    )

    if ($null -eq $Fields) {
        throw "fields is required."
    }

    $properties = @($Fields.PSObject.Properties)
    if ($properties.Count -eq 0) {
        throw "At least one Kanban field is required."
    }

    $namespaceManager = New-OneNoteNamespaceManager -Document $Document
    $namespaceUri = $Document.DocumentElement.NamespaceURI
    $outline = $null
    $updatedCount = 0

    foreach ($property in $properties) {
        $fieldName = [string]$property.Name
        $value = [string]$property.Value
        if ([string]::IsNullOrWhiteSpace($value)) {
            throw "Kanban field '$fieldName' must not be empty."
        }

        $fieldNode = Find-UniqueKanbanFieldNode `
            -Document $Document `
            -NamespaceManager $namespaceManager `
            -FieldName $fieldName
        $fieldOutline = Get-OneNoteOutlineAncestor -Node $fieldNode
        if ($null -eq $outline) {
            $outline = $fieldOutline
        }
        elseif ($outline -ne $fieldOutline) {
            throw "All Kanban fields must be inside the same Outline."
        }

        $label = Get-KanbanFieldLabel -FieldName $fieldName
        Set-OneNoteEntryText `
            -Document $Document `
            -EntryNode $fieldNode `
            -NamespaceManager $namespaceManager `
            -NamespaceUri $namespaceUri `
            -Text "$label : $value"
        $updatedCount++
    }

    return [pscustomobject]@{
        UpdatedCount = $updatedCount
        Outline = $outline
    }
}

function Update-KanbanSectionInDocument {
    param(
        [xml]$Document,
        [string]$Section,
        [string]$Mode,
        [string]$EntryId,
        [string[]]$Items
    )

    if ([string]::IsNullOrWhiteSpace($EntryId)) {
        throw "entryId is required."
    }
    if ($Mode -notin @("append", "replace")) {
        throw "mode must be 'append' or 'replace'."
    }

    $validItems = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($validItems.Count -eq 0) {
        throw "At least one non-empty section item is required."
    }

    $namespaceManager = New-OneNoteNamespaceManager -Document $Document
    $heading = Find-UniqueKanbanSectionHeading `
        -Document $Document `
        -NamespaceManager $namespaceManager `
        -Section $Section

    if (Test-OneNoteEntryMarker -RootNode $heading -NamespaceManager $namespaceManager -EntryId $EntryId) {
        return [pscustomobject]@{
            UpdatedCount = 0
            Duplicate = $true
            Outline = Get-OneNoteOutlineAncestor -Node $heading
        }
    }

    $namespaceUri = $Document.DocumentElement.NamespaceURI
    $children = $heading.SelectSingleNode("./one:OEChildren", $namespaceManager)
    if ($null -eq $children) {
        $children = $Document.CreateElement("one", "OEChildren", $namespaceUri)
        [void]$heading.AppendChild($children)
    }

    $existingItems = @($children.SelectNodes("./one:OE", $namespaceManager))
    $templateItem = $null
    if ($existingItems.Count -gt 0) {
        $templateItem = $existingItems[$existingItems.Count - 1]
    }

    $blankItem = $null
    if ($Mode -eq "replace") {
        foreach ($existingItem in $existingItems) {
            [void]$children.RemoveChild($existingItem)
        }
    }
    else {
        foreach ($existingItem in $existingItems) {
            if ([string]::IsNullOrWhiteSpace((Get-OneNoteDirectPlainText -Node $existingItem -NamespaceManager $namespaceManager))) {
                $blankItem = $existingItem
                break
            }
        }
    }

    foreach ($itemText in $validItems) {
        if ($null -ne $blankItem) {
            Add-OneNoteEntryMarker `
                -Document $Document `
                -EntryNode $blankItem `
                -NamespaceUri $namespaceUri `
                -EntryId $EntryId
            Set-OneNoteEntryText `
                -Document $Document `
                -EntryNode $blankItem `
                -NamespaceManager $namespaceManager `
                -NamespaceUri $namespaceUri `
                -Text $itemText
            $templateItem = $blankItem
            $blankItem = $null
            continue
        }

        $newEntry = New-OneNoteWorklogEntryNode `
            -Document $Document `
            -TemplateNode $templateItem `
            -NamespaceUri $namespaceUri `
            -EntryId $EntryId `
            -Text $itemText
        [void]$children.AppendChild($newEntry)
        $templateItem = $newEntry
    }

    return [pscustomobject]@{
        UpdatedCount = $validItems.Count
        Duplicate = $false
        Outline = Get-OneNoteOutlineAncestor -Node $heading
    }
}

function Add-KanbanWorklogToDocument {
    param(
        [xml]$Document,
        [string]$EntryId,
        [string[]]$Entries
    )

    $result = Update-KanbanSectionInDocument `
        -Document $Document `
        -Section "work" `
        -Mode "append" `
        -EntryId $EntryId `
        -Items $Entries

    return [pscustomobject]@{
        AddedCount = $result.UpdatedCount
        Duplicate = $result.Duplicate
        Outline = $result.Outline
    }
}

function New-OneNoteOutlinePatch {
    param(
        [xml]$SourceDocument,
        [string]$PageId,
        [System.Xml.XmlNode]$Outline
    )

    if ([string]::IsNullOrWhiteSpace($PageId)) {
        throw "pageId is required."
    }
    if ($null -eq $Outline.Attributes["objectID"]) {
        throw "The target Outline has no OneNote objectID."
    }

    $namespaceUri = $SourceDocument.DocumentElement.NamespaceURI
    $patch = New-Object System.Xml.XmlDocument
    $pageNode = $patch.CreateElement("one", "Page", $namespaceUri)
    $pageNode.SetAttribute("ID", $PageId)
    [void]$patch.AppendChild($pageNode)
    [void]$pageNode.AppendChild($patch.ImportNode($Outline, $true))
    return $patch
}
