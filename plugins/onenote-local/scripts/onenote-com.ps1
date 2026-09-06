param(
    [Parameter(Mandatory = $true)]
    [string]$Operation,

    [string]$InputJson = "{}"
)

$ErrorActionPreference = "Stop"

Set-StrictMode -Version 2.0

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

function Get-InputObject {
    if ([string]::IsNullOrWhiteSpace($InputJson)) {
        return [pscustomobject]@{}
    }

    return $InputJson | ConvertFrom-Json
}

function Get-JsonProperty {
    param(
        [object]$Object,
        [string]$Name,
        [object]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function Write-Json {
    param([object]$Value)

    $Value | ConvertTo-Json -Depth 40 -Compress
}

function New-OneNoteApplication {
    try {
        return New-Object -ComObject OneNote.Application
    }
    catch {
        throw "Microsoft OneNote desktop COM automation is not available. Install or open OneNote desktop and try again. $($_.Exception.Message)"
    }
}

function Get-ScopeValue {
    param([string]$Scope)

    switch (($Scope | ForEach-Object { "$_".ToLowerInvariant() })) {
        "self" { return 0 }
        "children" { return 1 }
        "notebooks" { return 2 }
        "sections" { return 3 }
        "pages" { return 4 }
        default { return 2 }
    }
}

function Get-AttributeValue {
    param(
        [System.Xml.XmlNode]$Node,
        [string]$Name
    )

    $attribute = $Node.Attributes[$Name]
    if ($null -eq $attribute) {
        return $null
    }

    return $attribute.Value
}

function Convert-HierarchyNode {
    param([System.Xml.XmlNode]$Node)

    $children = @()
    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $children += Convert-HierarchyNode -Node $child
        }
    }

    return [pscustomobject]@{
        type = $Node.LocalName
        name = Get-AttributeValue -Node $Node -Name "name"
        id = Get-AttributeValue -Node $Node -Name "ID"
        path = Get-AttributeValue -Node $Node -Name "path"
        lastModifiedTime = Get-AttributeValue -Node $Node -Name "lastModifiedTime"
        children = $children
    }
}

function Get-Hierarchy {
    param(
        [object]$OneNote,
        [string]$StartNodeId,
        [string]$Scope
    )

    $xmlText = ""
    $OneNote.GetHierarchy($StartNodeId, (Get-ScopeValue -Scope $Scope), [ref]$xmlText)
    [xml]$xml = $xmlText

    $items = @()
    foreach ($node in $xml.DocumentElement.ChildNodes) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $items += Convert-HierarchyNode -Node $node
        }
    }

    return [pscustomobject]@{
        items = $items
    }
}

function Get-AllPages {
    param([object]$OneNote)

    $hierarchy = Get-Hierarchy -OneNote $OneNote -StartNodeId "" -Scope "pages"
    $pages = New-Object System.Collections.Generic.List[object]

    function Visit-Node {
        param([object]$Node)

        if ($Node.type -eq "Page") {
            $pages.Add([pscustomobject]@{
                name = $Node.name
                id = $Node.id
                path = $Node.path
                lastModifiedTime = $Node.lastModifiedTime
            })
        }

        foreach ($child in @($Node.children)) {
            Visit-Node -Node $child
        }
    }

    foreach ($item in @($hierarchy.items)) {
        Visit-Node -Node $item
    }

    return $pages
}

function Get-PageXml {
    param(
        [object]$OneNote,
        [string]$PageId
    )

    $xmlText = ""
    $OneNote.GetPageContent($PageId, [ref]$xmlText, 0)
    return $xmlText
}

function Get-PagePlainText {
    param([xml]$PageXml)

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($node in $PageXml.GetElementsByTagName("*")) {
        if ($node.LocalName -eq "T") {
            $text = "$($node.InnerText)".Trim()
            if ($text.Length -gt 0) {
                $parts.Add($text)
            }
        }
    }

    return ($parts -join [Environment]::NewLine)
}

function New-OneNoteTextNode {
    param(
        [xml]$Document,
        [string]$NamespaceUri,
        [string]$Text
    )

    $tNode = $Document.CreateElement("one", "T", $NamespaceUri)
    [void]$tNode.AppendChild($Document.CreateCDataSection($Text))
    return $tNode
}

function Set-PageTitle {
    param(
        [xml]$Document,
        [string]$NamespaceUri,
        [string]$Title
    )

    $titleNode = $null
    foreach ($node in $Document.DocumentElement.ChildNodes) {
        if ($node.LocalName -eq "Title") {
            $titleNode = $node
            break
        }
    }

    if ($null -eq $titleNode) {
        $titleNode = $Document.CreateElement("one", "Title", $NamespaceUri)
        [void]$Document.DocumentElement.PrependChild($titleNode)
    }

    $oeNode = $null
    foreach ($node in $titleNode.ChildNodes) {
        if ($node.LocalName -eq "OE") {
            $oeNode = $node
            break
        }
    }

    if ($null -eq $oeNode) {
        $oeNode = $Document.CreateElement("one", "OE", $NamespaceUri)
        [void]$titleNode.AppendChild($oeNode)
    }

    $existingText = $null
    foreach ($node in $oeNode.ChildNodes) {
        if ($node.LocalName -eq "T") {
            $existingText = $node
            break
        }
    }

    $newText = New-OneNoteTextNode -Document $Document -NamespaceUri $NamespaceUri -Text $Title
    if ($null -eq $existingText) {
        [void]$oeNode.AppendChild($newText)
    }
    else {
        [void]$oeNode.ReplaceChild($newText, $existingText)
    }
}

function Add-OutlineText {
    param(
        [xml]$Document,
        [string]$NamespaceUri,
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $outline = $Document.CreateElement("one", "Outline", $NamespaceUri)
    $position = $Document.CreateElement("one", "Position", $NamespaceUri)
    $position.SetAttribute("x", "36.0")
    $position.SetAttribute("y", "120.0")
    $position.SetAttribute("z", "0")
    [void]$outline.AppendChild($position)

    $oeChildren = $Document.CreateElement("one", "OEChildren", $NamespaceUri)
    $oe = $Document.CreateElement("one", "OE", $NamespaceUri)
    [void]$oe.AppendChild((New-OneNoteTextNode -Document $Document -NamespaceUri $NamespaceUri -Text $Text))
    [void]$oeChildren.AppendChild($oe)
    [void]$outline.AppendChild($oeChildren)
    [void]$Document.DocumentElement.AppendChild($outline)
}

$argsObject = Get-InputObject
$oneNote = New-OneNoteApplication

switch ($Operation) {
    "status" {
        $result = Get-Hierarchy -OneNote $oneNote -StartNodeId "" -Scope "notebooks"
        Write-Json ([pscustomobject]@{
            available = $true
            notebookCount = @($result.items).Count
        })
    }
    "list_hierarchy" {
        $startNodeId = [string](Get-JsonProperty -Object $argsObject -Name "startNodeId" -DefaultValue "")
        $scope = [string](Get-JsonProperty -Object $argsObject -Name "scope" -DefaultValue "notebooks")
        Write-Json (Get-Hierarchy -OneNote $oneNote -StartNodeId $startNodeId -Scope $scope)
    }
    "search_pages" {
        $query = [string](Get-JsonProperty -Object $argsObject -Name "query" -DefaultValue "")
        $maxResults = [int](Get-JsonProperty -Object $argsObject -Name "maxResults" -DefaultValue 25)
        if ([string]::IsNullOrWhiteSpace($query)) {
            throw "query is required."
        }

        $matches = @(
            Get-AllPages -OneNote $oneNote |
                Where-Object { $_.name -like "*$query*" } |
                Select-Object -First $maxResults
        )
        Write-Json ([pscustomobject]@{ pages = $matches })
    }
    "read_page" {
        $pageId = [string](Get-JsonProperty -Object $argsObject -Name "pageId" -DefaultValue "")
        $includeXml = [bool](Get-JsonProperty -Object $argsObject -Name "includeXml" -DefaultValue $false)
        if ([string]::IsNullOrWhiteSpace($pageId)) {
            throw "pageId is required."
        }

        $xmlText = Get-PageXml -OneNote $oneNote -PageId $pageId
        [xml]$xml = $xmlText
        $payload = [ordered]@{
            pageId = $pageId
            title = Get-AttributeValue -Node $xml.DocumentElement -Name "name"
            text = Get-PagePlainText -PageXml $xml
        }
        if ($includeXml) {
            $payload["xml"] = $xmlText
        }
        Write-Json ([pscustomobject]$payload)
    }
    "open_page" {
        $pageId = [string](Get-JsonProperty -Object $argsObject -Name "pageId" -DefaultValue "")
        if ([string]::IsNullOrWhiteSpace($pageId)) {
            throw "pageId is required."
        }

        $oneNote.NavigateTo($pageId)
        Write-Json ([pscustomobject]@{
            opened = $true
            pageId = $pageId
        })
    }
    "create_page" {
        $sectionId = [string](Get-JsonProperty -Object $argsObject -Name "sectionId" -DefaultValue "")
        $title = [string](Get-JsonProperty -Object $argsObject -Name "title" -DefaultValue "")
        $bodyText = [string](Get-JsonProperty -Object $argsObject -Name "bodyText" -DefaultValue "")
        $openAfterCreate = [bool](Get-JsonProperty -Object $argsObject -Name "openAfterCreate" -DefaultValue $false)
        if ([string]::IsNullOrWhiteSpace($sectionId)) {
            throw "sectionId is required."
        }
        if ([string]::IsNullOrWhiteSpace($title)) {
            throw "title is required."
        }

        $pageId = ""
        $oneNote.CreateNewPage($sectionId, [ref]$pageId, 0)
        $xmlText = Get-PageXml -OneNote $oneNote -PageId $pageId
        [xml]$xml = $xmlText
        $namespaceUri = $xml.DocumentElement.NamespaceURI
        Set-PageTitle -Document $xml -NamespaceUri $namespaceUri -Title $title
        Add-OutlineText -Document $xml -NamespaceUri $namespaceUri -Text $bodyText
        $oneNote.UpdatePageContent($xml.OuterXml)
        if ($openAfterCreate) {
            $oneNote.NavigateTo($pageId)
        }

        Write-Json ([pscustomobject]@{
            created = $true
            pageId = $pageId
            title = $title
        })
    }
    "append_text" {
        $pageId = [string](Get-JsonProperty -Object $argsObject -Name "pageId" -DefaultValue "")
        $text = [string](Get-JsonProperty -Object $argsObject -Name "text" -DefaultValue "")
        $openAfterAppend = [bool](Get-JsonProperty -Object $argsObject -Name "openAfterAppend" -DefaultValue $false)
        if ([string]::IsNullOrWhiteSpace($pageId)) {
            throw "pageId is required."
        }
        if ([string]::IsNullOrWhiteSpace($text)) {
            throw "text is required."
        }

        $xmlText = Get-PageXml -OneNote $oneNote -PageId $pageId
        [xml]$xml = $xmlText
        Add-OutlineText -Document $xml -NamespaceUri $xml.DocumentElement.NamespaceURI -Text $text
        $oneNote.UpdatePageContent($xml.OuterXml)
        if ($openAfterAppend) {
            $oneNote.NavigateTo($pageId)
        }

        Write-Json ([pscustomobject]@{
            appended = $true
            pageId = $pageId
        })
    }
    default {
        throw "Unsupported operation: $Operation"
    }
}
