# Author: Claire Rosario
# Date Created: 2025-08-13
# Version: 1.0.1
# Last Updated: 2025-08-13
# Description:
#  Scans only Notes folders and exports IPM.StickyNote items to individual Markdown or TXT files.
# How to run:
#  1) Open 64-bit Windows PowerShell.
#  2) cd to the folder where this script lives.
#  3) Unblock-File .\export-outlook-notes.ps1
#  4) .\export-outlook-notes.ps1

Set-StrictMode -Version Latest

function prompt-yesno([string]$message, [bool]$default=$true) {
	$def = $(if ($default) { "Y" } else { "N" })
	while ($true) {
		$resp = Read-Host "$message (Y/N) [$def]"
		if ([string]::IsNullOrWhiteSpace($resp)) { return $default }
		switch ($resp.Trim().ToUpper()) {
			"Y" { return $true }
			"YES" { return $true }
			"N" { return $false }
			"NO" { return $false }
			default { Write-Host "Enter Y or N." }
		}
	}
}

function prompt-format([string]$default="md") {
	while ($true) {
		$resp = Read-Host "Output format md or txt [$default]"
		if ([string]::IsNullOrWhiteSpace($resp)) { return $default }
		$resp = $resp.Trim().ToLower()
		if ($resp -eq "md" -or $resp -eq "txt") { return $resp }
		Write-Host "Enter md or txt."
	}
}

function prompt-int([string]$message, [int]$default) {
	while ($true) {
		$resp = Read-Host "$message [$default]"
		if ([string]::IsNullOrWhiteSpace($resp)) { return $default }
		if ([int]::TryParse($resp, [ref]$null)) { return [int]$resp }
		Write-Host "Enter a valid integer."
	}
}

function remove-diacritics([string]$s) {
	if (-not $s) { return $s }
	$normalized = $s.Normalize([Text.NormalizationForm]::FormD)
	$sb = New-Object System.Text.StringBuilder
	foreach ($ch in $normalized.ToCharArray()) {
		if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
			[void]$sb.Append($ch)
		}
	}
	return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function slugify([string]$s, [int]$max_len = 80) {
	if (-not $s) { return "note" }
	$s = remove-diacritics($s)
	$invalid = [regex]::Escape([string]::Join("", [IO.Path]::GetInvalidFileNameChars()))
	$s = [regex]::Replace($s, "[$invalid]", "-")
	# Also kill PowerShell wildcard tokens
	$s = $s -replace "[\[\]\*\?`]", "-"
	$s = ($s -replace "\s+", " ").Trim()
	$s = $s -replace "[\s\-]+","-"
	$s = $s.Trim(".-")
	if ($s.Length -gt $max_len) { $s = $s.Substring(0, $max_len).Trim(".-") }
	if (-not $s) { $s = "note" }
	return $s
}

function unique-path([string]$base_path) {
	if (-not (Test-Path -LiteralPath $base_path)) { return $base_path }
	$dir = Split-Path -Parent $base_path
	$name = Split-Path -Leaf $base_path
	$stem = [IO.Path]::GetFileNameWithoutExtension($name)
	$ext = [IO.Path]::GetExtension($name)
	$i = 1
	while ($true) {
		$candidate = Join-Path $dir ("{0}_{1}{2}" -f $stem, $i, $ext)
		if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
		$i++
	}
}

function get-store-label($store) {
	$label = $store.DisplayName
	$file_path = ""
	try { $file_path = [string]$store.FilePath } catch {}
	if ($file_path) {
		if ($file_path.ToLower().EndsWith(".pst")) { return "$label [PST]" }
		if ($file_path.ToLower().EndsWith(".ost")) { return "$label [OST]" }
	}
	return "$label [Mailbox]"
}

function collect-notes-folders($folder, [ref]$accum) {
	try {
		# 5 = olNoteItem
		if ($folder.DefaultItemType -eq 5) {
			$accum.Value.Add($folder) | Out-Null
		}
	} catch {}
	try {
		foreach ($sub in $folder.Folders) {
			collect-notes-folders -folder $sub -accum $accum
		}
	} catch {}
}

Write-Host "Outlook Notes fast exporter - interactive"

$format = prompt-format "md"
$out_dir_default = $(if ($format -eq "md") { ".\notes_md" } else { ".\notes_txt" })
$out_dir = Read-Host "Output folder [$out_dir_default]"
if ([string]::IsNullOrWhiteSpace($out_dir)) { $out_dir = $out_dir_default }
$include_index_csv = prompt-yesno "Write index CSV too" $false
$prefix_date = prompt-yesno "Prefix date in filenames (YYYYMMDD)" $false
$flatten = prompt-yesno "Flatten output into a single folder" $false
$max_name_len = prompt-int "Max filename length" 80
$dry_run = prompt-yesno "Dry run - list actions only" $false

if (-not (Test-Path -Path $out_dir)) { New-Item -Path $out_dir -ItemType Directory -Force | Out-Null }

$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")

$stores = @()
try { $stores = @($ns.Stores) } catch { $stores = @() }
if ($stores.Count -eq 0) { $stores = @($ns.GetDefaultFolder(12).Store) } # olFolderNotes

$index = New-Object System.Collections.Generic.List[object]
$note_count = 0

foreach ($store in $stores) {
	$store_label = get-store-label $store
	$root = $null
	try { $root = $store.GetRootFolder() } catch { continue }

	$notes_folders = New-Object System.Collections.Generic.List[object]
	collect-notes-folders -folder $root -accum ([ref]$notes_folders)

	foreach ($nf in $notes_folders) {
		$folder_path = ""
		try { $folder_path = $nf.FolderPath } catch { $folder_path = $nf.Name }

		$items = $null
		try { $items = $nf.Items } catch { continue }

		# Filter to sticky notes only
		$filtered = $null
		try { $filtered = $items.Restrict("[MessageClass] Like 'IPM.StickyNote%'") } catch { $filtered = $items }

$total_notes = $filtered.Count
$note_count = 0
		foreach ($item in $filtered) {
			if ($null -eq $item) { continue }
			$msg_class = ""
			try { $msg_class = [string]$item.MessageClass } catch {}
			if ($msg_class -notlike "IPM.StickyNote*") { continue }

			$subject = ""
			$body = ""
			$categories = ""
			$color = ""
			$created = $null
			$modified = $null
			$entry_id = ""

			try { $subject = ($item.Subject | Out-String).Trim() } catch {}
			try { $body = ($item.Body | Out-String).Replace("`r`n","`n").Trim() } catch {}
			try { $categories = ($item.Categories | Out-String).Trim() } catch {}
			try { $color = [string]$item.Color } catch {}
			try { $created = $item.CreationTime } catch {}
			try { $modified = $item.LastModificationTime } catch {}
			try { $entry_id = [string]$item.EntryID } catch {}

			if ([string]::IsNullOrWhiteSpace($subject)) {
				if (-not [string]::IsNullOrWhiteSpace($body)) {
					$subject = ($body -split "`n")[0].Trim()
				}
				if ([string]::IsNullOrWhiteSpace($subject)) { $subject = "(no subject)" }
			}

			$note_count++
		Write-Progress -Activity "Exporting Outlook Notes" -Status ("Processing {0} of {1}" -f $note_count, $total_notes) -PercentComplete (($note_count / $total_notes) * 100)
		Write-Host ("[{0}/{1}] " -f $note_count, $total_notes) -NoNewline; Write-Host $subject
			if ($dry_run) { continue }

			$date_prefix = ""
			if ($prefix_date -and $created) { try { $date_prefix = (Get-Date $created -Format "yyyyMMdd") + "-" } catch {} }
			$slug = slugify $subject $max_name_len
			$ext = $(if ($format -eq "md") { ".md" } else { ".txt" })
			$file_name = "$date_prefix$slug$ext"

			$target_dir = $out_dir
			if (-not $flatten) {
				$store_dir = Join-Path $target_dir (slugify $store_label 64)
				if (-not (Test-Path -Path $store_dir)) { New-Item -Path $store_dir -ItemType Directory -Force | Out-Null }
				$target_dir = $store_dir
			}

			$out_path = Join-Path $target_dir $file_name
			$out_path = unique-path $out_path

			if ($format -eq "md") {
				$yaml = @("---")
				$yaml += "title: ""$subject"""
				$yaml += "source: Outlook Notes"
				if ($created) { $yaml += "created: ""$((Get-Date $created -Format o))""" }
				if ($modified) { $yaml += "modified: ""$((Get-Date $modified -Format o))""" }
				if ($categories) { $yaml += "categories: ""$categories""" }
				if ($color) { $yaml += "color: ""$color""" }
				$yaml += "store: ""$store_label"""
				$yaml += "folder: ""$folder_path"""
				if ($entry_id) { $yaml += "entry_id: ""$entry_id""" }
				$yaml += @("---","")
				$content = @($yaml -join "`r`n", $body) -join "`r`n"
			} else {
				$hdr = @(
					"Title: $subject",
					"Source: Outlook Notes"
				)
				if ($created) { $hdr += "Created: $((Get-Date $created -Format o))" }
				if ($modified) { $hdr += "Modified: $((Get-Date $modified -Format o))" }
				if ($categories) { $hdr += "Categories: $categories" }
				if ($color) { $hdr += "Color: $color" }
				$hdr += "Store: $store_label"
				$hdr += "Folder: $folder_path"
				if ($entry_id) { $hdr += "EntryID: $entry_id" }
				$content = @($hdr -join "`r`n","---","", $body) -join "`r`n"
			}

			if (-not (Test-Path -LiteralPath (Split-Path $out_path))) { New-Item -Path (Split-Path $out_path) -ItemType Directory -Force | Out-Null }
			Set-Content -LiteralPath $out_path -Value $content -Encoding utf8

			if ($include_index_csv) {
				$index.Add([pscustomobject]@{
					StoreLabel       = $store_label
					FolderPath       = $folder_path
					Subject          = $subject
					Body             = $body
					Categories       = $categories
					Color            = $color
					CreationTime     = $(if ($created) { (Get-Date $created -Format o) } else { "" })
					LastModification = $(if ($modified) { (Get-Date $modified -Format o) } else { "" })
					EntryID          = $entry_id
				}) | Out-Null
			}
		}
	}
}

if ($dry_run) {
	Write-Host "Dry run complete. Would write $note_count files."
	exit 0
}

if ($include_index_csv) {
	$index_path = Join-Path $out_dir "outlook_notes_index.csv"
	$index | Export-Csv -Path $index_path -Encoding UTF8 -NoTypeInformation
	Write-Host "Index CSV: $index_path"
}

Write-Host "Export complete. Wrote $note_count files."
Write-Host "Export complete: $note_count notes saved." -ForegroundColor Green