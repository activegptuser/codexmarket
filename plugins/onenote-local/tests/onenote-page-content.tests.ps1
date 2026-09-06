$ErrorActionPreference = "Stop"

. "$PSScriptRoot\..\scripts\onenote-page-content.ps1"

function Assert-Equal {
    param(
        [object]$Expected,
        [object]$Actual,
        [string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "$Message Expected=[$Expected] Actual=[$Actual]"
    }
}

function ConvertFrom-CodePoints {
    param([int[]]$CodePoints)

    return -join @($CodePoints | ForEach-Object { [char]$_ })
}

$workHeadingText = (ConvertFrom-CodePoints @(0xC791, 0xC5C5)) + "(" +
    (ConvertFrom-CodePoints @(0xC138, 0xBD80)) + " " +
    (ConvertFrom-CodePoints @(0xB0B4, 0xC6A9)) + ")"
$doneHeadingText = ConvertFrom-CodePoints @(0xC644, 0xB8CC, 0xC870, 0xAC74)
$existingDoneText = (ConvertFrom-CodePoints @(0xAE30, 0xC874)) + " " + $doneHeadingText
$assigneeLabel = ConvertFrom-CodePoints @(0xB2F4, 0xB2F9)
$dueDateLabel = ConvertFrom-CodePoints @(0xB9C8, 0xAC10, 0xC77C)
$prerequisitesHeadingText = ConvertFrom-CodePoints @(0xC120, 0xD589, 0xC5C5, 0xBB34)
$risksHeadingText = (ConvertFrom-CodePoints @(0xB9AC, 0xC2A4, 0xD06C)) + " / " +
    (ConvertFrom-CodePoints @(0xC774, 0xC288))

function New-TestPageDocument {
    $xmlText = @'
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="{PAGE-ID}">
  <one:Outline objectID="{OUTLINE-ID}">
    <one:Position x="36.0" y="104.4" z="0" />
    <one:OEChildren>
      <one:OE objectID="{ASSIGNEE-ID}"><one:T><![CDATA[__ASSIGNEE_LABEL__ : Old owner]]></one:T></one:OE>
      <one:OE objectID="{DUE-DATE-ID}"><one:T><![CDATA[__DUE_DATE_LABEL__ : -]]></one:T></one:OE>
      <one:OE objectID="{WORK-HEADING-ID}">
        <one:T><![CDATA[<span>__WORK_HEADING__</span>]]></one:T>
        <one:OEChildren>
          <one:OE objectID="{EMPTY-ITEM-ID}">
            <one:List><one:Bullet bullet="25" fontSize="11.0" /></one:List>
            <one:T><![CDATA[]]></one:T>
          </one:OE>
        </one:OEChildren>
      </one:OE>
      <one:OE objectID="{DONE-HEADING-ID}">
        <one:T><![CDATA[<span>__DONE_HEADING__</span>]]></one:T>
        <one:OEChildren>
          <one:OE objectID="{DONE-ITEM-ID}">
            <one:List><one:Bullet bullet="25" fontSize="11.0" /></one:List>
            <one:T><![CDATA[__DONE_TEXT__]]></one:T>
          </one:OE>
        </one:OEChildren>
      </one:OE>
      <one:OE objectID="{PREREQUISITES-HEADING-ID}">
        <one:T><![CDATA[<span>__PREREQUISITES_HEADING__</span>]]></one:T>
        <one:OEChildren>
          <one:OE objectID="{PREREQUISITES-ITEM-ID}">
            <one:List><one:Bullet bullet="25" fontSize="11.0" /></one:List>
            <one:T><![CDATA[Existing prerequisite]]></one:T>
          </one:OE>
        </one:OEChildren>
      </one:OE>
      <one:OE objectID="{RISKS-HEADING-ID}">
        <one:T><![CDATA[<span>__RISKS_HEADING__</span>]]></one:T>
        <one:OEChildren>
          <one:OE objectID="{RISKS-ITEM-ID}">
            <one:List><one:Bullet bullet="25" fontSize="11.0" /></one:List>
            <one:T><![CDATA[Old risk]]></one:T>
          </one:OE>
        </one:OEChildren>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>
'@

    $xmlText = $xmlText.Replace("__WORK_HEADING__", $workHeadingText)
    $xmlText = $xmlText.Replace("__DONE_HEADING__", $doneHeadingText)
    $xmlText = $xmlText.Replace("__DONE_TEXT__", $existingDoneText)
    $xmlText = $xmlText.Replace("__ASSIGNEE_LABEL__", $assigneeLabel)
    $xmlText = $xmlText.Replace("__DUE_DATE_LABEL__", $dueDateLabel)
    $xmlText = $xmlText.Replace("__PREREQUISITES_HEADING__", $prerequisitesHeadingText)
    $xmlText = $xmlText.Replace("__RISKS_HEADING__", $risksHeadingText)
    return [xml]$xmlText
}

$fieldDocument = New-TestPageDocument
$fieldResult = Update-KanbanFieldsInDocument `
    -Document $fieldDocument `
    -Fields ([pscustomobject]@{
        assignee = "New owner"
        dueDate = "2026-09-06"
    })
$fieldNamespaceManager = New-OneNoteNamespaceManager -Document $fieldDocument
$fieldTexts = @($fieldDocument.SelectNodes("//one:OE/one:T", $fieldNamespaceManager) | ForEach-Object {
    ConvertFrom-OneNoteHtmlText -Text $_.InnerText
})

Assert-Equal 2 $fieldResult.UpdatedCount "Two scalar fields should be updated."
Assert-Equal $true ($fieldTexts -contains "$assigneeLabel : New owner") "The assignee should be updated."
Assert-Equal $true ($fieldTexts -contains "$dueDateLabel : 2026-09-06") "The due date should be updated."
Assert-Equal $true ($fieldTexts -contains $existingDoneText) "Updating scalar fields must not change section items."

$sectionDocument = New-TestPageDocument
$completionResult = Update-KanbanSectionInDocument `
    -Document $sectionDocument `
    -Section "completion" `
    -Mode "append" `
    -EntryId "completion-20260906" `
    -Items @("New completion item")

Assert-Equal 1 $completionResult.UpdatedCount "One completion item should be appended."
Assert-Equal $false $completionResult.Duplicate "The first section update should not be a duplicate."

$sectionNamespaceManager = New-OneNoteNamespaceManager -Document $sectionDocument
$completionHeading = Find-UniqueKanbanSectionHeading `
    -Document $sectionDocument `
    -NamespaceManager $sectionNamespaceManager `
    -Section "completion"
$completionItems = @($completionHeading.SelectNodes("./one:OEChildren/one:OE", $sectionNamespaceManager))
Assert-Equal 2 $completionItems.Count "Append mode must preserve the existing completion item."
Assert-Equal "New completion item" (Get-OneNoteNodePlainText -Node $completionItems[1]) "The completion item differs."

$duplicateSectionResult = Update-KanbanSectionInDocument `
    -Document $sectionDocument `
    -Section "completion" `
    -Mode "append" `
    -EntryId "completion-20260906" `
    -Items @("Must not be added")
Assert-Equal $true $duplicateSectionResult.Duplicate "A duplicate section update must be reported."
Assert-Equal 2 @($completionHeading.SelectNodes("./one:OEChildren/one:OE", $sectionNamespaceManager)).Count "A duplicate section update must not add items."

$replaceResult = Update-KanbanSectionInDocument `
    -Document $sectionDocument `
    -Section "risks" `
    -Mode "replace" `
    -EntryId "risks-20260906" `
    -Items @("New risk 1", "New risk 2")
$risksHeading = Find-UniqueKanbanSectionHeading `
    -Document $sectionDocument `
    -NamespaceManager $sectionNamespaceManager `
    -Section "risks"
$riskItems = @($risksHeading.SelectNodes("./one:OEChildren/one:OE", $sectionNamespaceManager))
Assert-Equal 2 $replaceResult.UpdatedCount "Replace mode should add both replacement items."
Assert-Equal 2 $riskItems.Count "Replace mode must remove the old risk item only."
Assert-Equal "New risk 1" (Get-OneNoteNodePlainText -Node $riskItems[0]) "The first replacement risk differs."
Assert-Equal "Existing prerequisite" (Get-OneNoteNodePlainText -Node $sectionDocument.SelectSingleNode("//one:OE[@objectID='{PREREQUISITES-ITEM-ID}']", $sectionNamespaceManager)) "Replace mode must preserve other sections."

$document = New-TestPageDocument
$result = Add-KanbanWorklogToDocument `
    -Document $document `
    -EntryId "worklog-20260906" `
    -Entries @("MCP connection verified", "Local COM notebooks listed")

Assert-Equal 2 $result.AddedCount "Two worklog items should be added."
Assert-Equal $false $result.Duplicate "The first call should not be a duplicate."

$namespaceManager = New-OneNoteNamespaceManager -Document $document
$workHeading = Find-UniqueKanbanWorkHeading `
    -Document $document `
    -NamespaceManager $namespaceManager
$workItems = @($workHeading.SelectNodes("./one:OEChildren/one:OE", $namespaceManager))

Assert-Equal 2 $workItems.Count "The empty bullet should be reused before appending."
Assert-Equal "MCP connection verified" (Get-OneNoteNodePlainText -Node $workItems[0]) "The first worklog item differs."
Assert-Equal "Local COM notebooks listed" (Get-OneNoteNodePlainText -Node $workItems[1]) "The second worklog item differs."
Assert-Equal "worklog-20260906" $workItems[0].SelectSingleNode("./one:Meta[@name='onenote-local.entry-id']", $namespaceManager).GetAttribute("content") "The first marker is missing."
Assert-Equal "worklog-20260906" $workItems[1].SelectSingleNode("./one:Meta[@name='onenote-local.entry-id']", $namespaceManager).GetAttribute("content") "The second marker is missing."

$duplicateResult = Add-KanbanWorklogToDocument `
    -Document $document `
    -EntryId "worklog-20260906" `
    -Entries @("Must not be added again")

Assert-Equal 0 $duplicateResult.AddedCount "A duplicate call must not add items."
Assert-Equal $true $duplicateResult.Duplicate "A duplicate call must be reported."
Assert-Equal 2 @($workHeading.SelectNodes("./one:OEChildren/one:OE", $namespaceManager)).Count "A duplicate call must not change the item count."

$doneItem = $document.SelectSingleNode("//one:OE[one:T[contains(., '$doneHeadingText')]]/one:OEChildren/one:OE", $namespaceManager)
Assert-Equal $existingDoneText (Get-OneNoteNodePlainText -Node $doneItem) "An unrelated section must remain unchanged."

$missingDocument = New-TestPageDocument
$missingNamespaceManager = New-OneNoteNamespaceManager -Document $missingDocument
$missingHeadingNode = Find-UniqueKanbanWorkHeading -Document $missingDocument -NamespaceManager $missingNamespaceManager
$missingHeadingTextNode = $missingHeadingNode.SelectSingleNode("./one:T", $missingNamespaceManager)
$missingHeadingTextNode.RemoveAll()
[void]$missingHeadingTextNode.AppendChild($missingDocument.CreateCDataSection("Different heading"))

$missingFailed = $false
try {
    Add-KanbanWorklogToDocument `
        -Document $missingDocument `
        -EntryId "missing" `
        -Entries @("Must not be added")
}
catch {
    $missingFailed = $_.Exception.Message -like "*Kanban 'work' heading*"
}

Assert-Equal $true $missingFailed "A missing target heading must fail clearly."

$patch = New-OneNoteOutlinePatch `
    -SourceDocument $document `
    -PageId "{PAGE-ID}" `
    -Outline $result.Outline
$patchNamespaceManager = New-OneNoteNamespaceManager -Document $patch
Assert-Equal 1 @($patch.SelectNodes("/one:Page/one:Outline", $patchNamespaceManager)).Count "The update patch must contain exactly one Outline."
Assert-Equal "{OUTLINE-ID}" $patch.DocumentElement.FirstChild.GetAttribute("objectID") "The update patch must retain the OneNote Outline objectID."

Write-Output "onenote-page-content tests passed"
