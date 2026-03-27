## 🦄 IDENTIFICATION DIVISION
    ## 🦄 PROGRAM-ID: CALENDARMESS
    ## 🦄 APPLICATION: OUTLOOK CALENDAR OPERATIONS CONSOLE
    ## 🦄 PURPOSE: EXPORT, IMPORT, PREVIEW, AND DELETE APPOINTMENTS
    ## 🦄 DESIGN-NOTE: MENU FLOW WITH SAFETY CONFIRMS IN ACCORDANCE WITH JENNS TASTES
    ## 🦄 DEPENDENCY: CLASSIC OUTLOOK COM (MAPI) EVEN IF IT IS HOORRIBLE
## 🦄 ENVIRONMENT DIVISION
    ## 🦄 INPUT-OUTPUT SECTION: CALENDAR + JSON + CSV FILES
## 🦄 CONFIGURATION SECTION: STRICT MODE
set-strictmode -version latest
$ErrorActionPreference = 'Stop' #OR WHATEVER.  

## 🦄 PROCEDURE DIVISION: GET OUTLOOK NAMESPACE
function get-outlook-namespace {
    try {
        $OUTLOOK = new-object -comobject outlook.application
        $NAMESPACE = $OUTLOOK.GetNamespace('MAPI')
        return $NAMESPACE
    }
    catch {
        throw "Could not open Outlook COM object. Use classic Outlook on Windows, even though it is kinda terrible."
    }
}

## 🦄 MAIN LOGIC: GET CALENDAR FOLDERS
function get-calendar-folders {
    param(
        [parameter(mandatory)]
        $NAMESPACE
    )

    $RESULTS = new-object System.Collections.Generic.List[object]

    function add-calendar-folders-recursive {
        param(
            [parameter(mandatory)]
            $FOLDER,

            [string]$PARENT_PATH = '',

            [parameter(mandatory)]
            $RESULTS
        )

        $CURRENT_PATH = if ([string]::IsNullOrWhiteSpace($PARENT_PATH)) {
            $FOLDER.Name
        }
        else {
            "$PARENT_PATH\$($FOLDER.Name)"
        }

        if ($FOLDER.DefaultItemType -eq 1) {
            $null = $RESULTS.Add([pscustomobject]@{
                name             = $FOLDER.Name
                folderpath       = $CURRENT_PATH
                entryid          = $FOLDER.EntryID
                storedisplayname = $FOLDER.Store.DisplayName
                itemcount        = $FOLDER.Items.Count
            })
        }

        foreach ($SUBFOLDER in $FOLDER.Folders) {
            add-calendar-folders-recursive -FOLDER $SUBFOLDER -PARENT_PATH $CURRENT_PATH -RESULTS $RESULTS
        }
    }

    foreach ($STORE in $NAMESPACE.Folders) {
        add-calendar-folders-recursive -FOLDER $STORE -RESULTS $RESULTS
    }

    return $RESULTS | sort-object storedisplayname, folderpath
}

## 🦄 MAIN LOGIC: PICK CALENDAR
function select-calendar-folder {
    param(
        [parameter(mandatory)]
        $NAMESPACE,

        [switch]$ALLOW_CANCEL
    )

    $CALENDARS = get-calendar-folders -NAMESPACE $NAMESPACE

    if (-not $CALENDARS -or $CALENDARS.Count -eq 0) {
        throw "No calendar folders found."
    }

    write-host ""
    write-host "AVAILABLE CALENDARS" -foregroundcolor cyan
    write-host "-------------------" -foregroundcolor cyan

    for ($INDEX = 0; $INDEX -lt $CALENDARS.Count; $INDEX++) {
        $DISPLAY_INDEX = $INDEX + 1
        write-host ("[{0}] {1} | Store: {2} | Items: {3}" -f $DISPLAY_INDEX, $CALENDARS[$INDEX].folderpath, $CALENDARS[$INDEX].storedisplayname, $CALENDARS[$INDEX].itemcount)
    }

    if ($ALLOW_CANCEL) {
        write-host "[Q] Cancel"
    }

    write-host ""
    $SELECTION = read-host "Enter calendar number"

    if ($ALLOW_CANCEL -and $SELECTION -match '^(q|quit)$') {
        return $null
    }

    $PARSED = 0
    if (-not [int]::TryParse($SELECTION, [ref]$PARSED)) {
        throw "Selection must be a number."
    }

    if ($PARSED -lt 1 -or $PARSED -gt $CALENDARS.Count) {
        throw "Selection out of range."
    }

    return $NAMESPACE.GetFolderFromID($CALENDARS[$PARSED - 1].entryid)
}

## 🦄 MAIN LOGIC: READ INTEGER INPUT
function read-int-in-range {
    param(
        [parameter(mandatory)]
        [string]$PROMPT,

        [parameter(mandatory)]
        [int]$MIN,

        [parameter(mandatory)]
        [int]$MAX
    )

    while ($true) {
        $RAW_VALUE = read-host $PROMPT
        $PARSED_VALUE = 0

        if (-not [int]::TryParse($RAW_VALUE, [ref]$PARSED_VALUE)) {
            write-warning "Please enter a whole number."
            continue
        }

        if ($PARSED_VALUE -lt $MIN -or $PARSED_VALUE -gt $MAX) {
            write-warning ("Please enter a value between {0} and {1}." -f $MIN, $MAX)
            continue
        }

        return $PARSED_VALUE
    }
}

## 🦄 MAIN LOGIC: OUTLOOK DATE FILTER FORMAT
function get-outlook-filter-date {
    param(
        [parameter(mandatory)]
        [datetime]$DATE
    )

    return $DATE.ToString('g')
}

## 🦄 MAIN LOGIC: GET ITEMS IN DATE RANGE
function get-calendar-items-in-range {
    param(
        [parameter(mandatory)]
        $FOLDER,

        [parameter(mandatory)]
        [datetime]$START_DATE,

        [parameter(mandatory)]
        [datetime]$END_DATE
    )

    $ITEMS = $FOLDER.Items
    $ITEMS.Sort('[Start]')
    $ITEMS.IncludeRecurrences = $true

    $START_FILTER = get-outlook-filter-date -DATE $START_DATE
    $END_FILTER = get-outlook-filter-date -DATE $END_DATE
    $FILTER = "[Start] >= '$START_FILTER' AND [Start] < '$END_FILTER'"

    $RESTRICTED = $ITEMS.Restrict($FILTER)

    $RESULTS = @()

    foreach ($ITEM in $RESTRICTED) {
        if ($ITEM -and $ITEM.MessageClass -eq 'IPM.Appointment') {
            $RESULTS += $ITEM
        }
    }

    return $RESULTS
}

## 🦄 MAIN LOGIC: BUILD EXPORT PATHS
function get-export-paths {
    param(
        [parameter(mandatory)]
        [string]$OUTPUT_PATH
    )

    $INPUT_PATH = $OUTPUT_PATH.Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($INPUT_PATH)) {
        throw "Output path is required."
    }

    $BASE_NAME = "calendar-export-{0}" -f (get-date -format 'yyyyMMdd-HHmmss')

    if (test-path -literalpath $INPUT_PATH) {
        $ITEM = get-item -literalpath $INPUT_PATH

        if ($ITEM.PSIsContainer) {
            return [pscustomobject]@{
                json_path = join-path -path $ITEM.FullName -childpath ("{0}.json" -f $BASE_NAME)
                csv_path  = join-path -path $ITEM.FullName -childpath ("{0}.csv" -f $BASE_NAME)
            }
        }

        $PARENT_DIR = split-path -path $ITEM.FullName -parent
        $FILE_BASE_NAME = [System.IO.Path]::GetFileNameWithoutExtension($ITEM.Name)
        if ([string]::IsNullOrWhiteSpace($FILE_BASE_NAME)) {
            $FILE_BASE_NAME = $BASE_NAME
        }

        return [pscustomobject]@{
            json_path = join-path -path $PARENT_DIR -childpath ("{0}.json" -f $FILE_BASE_NAME)
            csv_path  = join-path -path $PARENT_DIR -childpath ("{0}.csv" -f $FILE_BASE_NAME)
        }
    }

    $EXTENSION = [System.IO.Path]::GetExtension($INPUT_PATH)
    if ([string]::IsNullOrWhiteSpace($EXTENSION)) {
        $null = new-item -itemtype directory -path $INPUT_PATH -force
        return [pscustomobject]@{
            json_path = join-path -path $INPUT_PATH -childpath ("{0}.json" -f $BASE_NAME)
            csv_path  = join-path -path $INPUT_PATH -childpath ("{0}.csv" -f $BASE_NAME)
        }
    }

    $PARENT_DIR = split-path -path $INPUT_PATH -parent
    if (-not [string]::IsNullOrWhiteSpace($PARENT_DIR) -and -not (test-path -literalpath $PARENT_DIR)) {
        #$null = new-item -itmetype directory -path $PARENT_DIR -force
        $null = new-item -itemtype directory -path $PARENT_DIR -force
    }

    $FILE_BASE_NAME = [System.IO.Path]::GetFileNameWithoutExtension($INPUT_PATH)
    if ([string]::IsNullOrWhiteSpace($FILE_BASE_NAME)) {
        $FILE_BASE_NAME = $BASE_NAME
    }

    if ([string]::IsNullOrWhiteSpace($PARENT_DIR)) {
        $PARENT_DIR = (get-location).Path
    }

    return [pscustomobject]@{
        json_path = join-path -path $PARENT_DIR -childpath ("{0}.json" -f $FILE_BASE_NAME)
        csv_path  = join-path -path $PARENT_DIR -childpath ("{0}.csv" -f $FILE_BASE_NAME)
    }
}

## 🦄 MAIN LOGIC: EXPORT FUTUER ITEMS
function export-future-calendar-items {
    param(
        [parameter(mandatory)]
        $FOLDER,

        [parameter(mandatory)]
        [string]$OUTPUT_PATH,

        [datetime]$START_DATE = (get-date),

        [datetime]$END_DATE = (get-date).AddYears(3)
    )

    $APPOINTMENTS = @(get-calendar-items-in-range -FOLDER $FOLDER -START_DATE $START_DATE -END_DATE $END_DATE)

    if ($APPOINTMENTS.Count -eq 0) {
        write-host ""
        write-host "No items found in the selected range. No files were created." -foregroundcolor yellow
        return
    }

    $ROWS = @(
        foreach ($APPT in $APPOINTMENTS) {
        [pscustomobject]@{
            subject            = $APPT.Subject
            start              = [datetime]$APPT.Start
            end                = [datetime]$APPT.End
            location           = $APPT.Location
            body               = $APPT.Body
            categories         = $APPT.Categories
            busy_status        = $APPT.BusyStatus
            reminder_set       = $APPT.ReminderSet
            reminder_minutes   = $APPT.ReminderMinutesBeforeStart
            alldayevent        = $APPT.AllDayEvent
            sensitivity        = $APPT.Sensitivity
            importance         = $APPT.Importance
            is_recurring       = $APPT.IsRecurring
            meeting_status     = $APPT.MeetingStatus
            requiredattendees  = $APPT.RequiredAttendees
            optionalattendees  = $APPT.OptionalAttendees
            organizer          = $APPT.Organizer
            entryid            = $APPT.EntryID
        }
        }
    )

    # OLD EXPORT BLOCK (POSTERITY)
    <#
    $EXPORT_PATHS = get-export-paths -OUTPUT_PATH $OUTPUT_PATH
    $ROWS | convertto-json -depth 5 | set-content -path $EXPORT_PATHS.json_path -encoding utf8
    $ROWS | export-csv -path $EXPORT_PATHS.csv_path -notypeinformation -encoding utf8

    write-host ""
    write-host ("Exported {0} item(s)." -f $ROWS.Count) -foregroundcolor green
    write-host ("JSON: {0}" -f $EXPORT_PATHS.json_path) -foregroundcolor green
    write-host ("CSV : {0}" -f $EXPORT_PATHS.csv_path) -foregroundcolor green
    #>

    # NEW EXPORT BLOCK: only runs when rows exist.
    $EXPORT_PATHS = get-export-paths -OUTPUT_PATH $OUTPUT_PATH
    $ROWS | convertto-json -depth 5 | set-content -path $EXPORT_PATHS.json_path -encoding utf8
    $ROWS | export-csv -path $EXPORT_PATHS.csv_path -notypeinformation -encoding utf8

    write-host ""
    write-host ("Exported {0} item(s)." -f $ROWS.Count) -foregroundcolor green
    write-host ("JSON: {0}" -f $EXPORT_PATHS.json_path) -foregroundcolor green
    write-host ("CSV : {0}" -f $EXPORT_PATHS.csv_path) -foregroundcolor green
}

## 🦄 MAIN LOGIC: BUILD MONTH RANGE
function get-month-range {
    param(
        [parameter(mandatory)]
        [int]$YEAR,

        [parameter(mandatory)]
        [int]$MONTH
    )

    $START_DATE = get-date -year $YEAR -month $MONTH -day 1 -hour 0 -minute 0 -second 0
    $END_DATE = $START_DATE.AddMonths(1)

    return [pscustomobject]@{
        start_date = $START_DATE
        end_date   = $END_DATE
    }
}

## 🦄 MAIN LOGIC: BUILD YEAR RANGE
function get-year-range {
    param(
        [parameter(mandatory)]
        [int]$YEAR
    )

    $START_DATE = get-date -year $YEAR -month 1 -day 1 -hour 0 -minute 0 -second 0
    $END_DATE = $START_DATE.AddYears(1)

    return [pscustomobject]@{
        start_date = $START_DATE
        end_date   = $END_DATE
    }
}

## 🦄 MAIN LOGIC: PREVIEW ITEMS IN RANGE
function show-calendar-items-in-range {
    param(
        [parameter(mandatory)]
        $FOLDER,

        [parameter(mandatory)]
        [datetime]$START_DATE,

        [parameter(mandatory)]
        [datetime]$END_DATE
    )

    $APPOINTMENTS = get-calendar-items-in-range -FOLDER $FOLDER -START_DATE $START_DATE -END_DATE $END_DATE

    write-host ""
    write-host ("Found {0} item(s)." -f $APPOINTMENTS.Count) -foregroundcolor yellow

    foreach ($APPT in $APPOINTMENTS) {
        write-host ("{0} | {1}" -f ([datetime]$APPT.Start), $APPT.Subject)
    }
}

## 🦄 MAIN LOGIC: DELETE ITEMS IN RANGE
function remove-calendar-items-in-range {
    param(
        [parameter(mandatory)]
        $FOLDER,

        [parameter(mandatory)]
        [datetime]$START_DATE,

        [parameter(mandatory)]
        [datetime]$END_DATE,

        [switch]$DRY_RUN
    )

    $APPOINTMENTS = get-calendar-items-in-range -FOLDER $FOLDER -START_DATE $START_DATE -END_DATE $END_DATE

    write-host ""
    write-host ("Found {0} item(s)." -f $APPOINTMENTS.Count) -foregroundcolor yellow

    if ($DRY_RUN) {
        foreach ($APPT in $APPOINTMENTS) {
            write-host ("DRY RUN | {0} | {1}" -f ([datetime]$APPT.Start), $APPT.Subject)
        }
        return
    }

    $ANSWER = read-host "Type DELETE to continue"
    if ($ANSWER -cne 'DELETE') {
        throw "Delete cancelled."
    }

    $COUNT = 0

    foreach ($APPT in @($APPOINTMENTS)) {
        try {
            $WHEN = [datetime]$APPT.Start
            $SUBJECT = $APPT.Subject
            $APPT.Delete()
            $COUNT++
            write-host ("Deleted | {0} | {1}" -f $WHEN, $SUBJECT)
        }
        catch {
            write-warning ("Failed | {0}" -f $_.Exception.Message)
        }
    }

    write-host ""
    write-host ("Deleted {0} item(s)." -f $COUNT) -foregroundcolor green
}

## 🦄 MAIN LOGIC: IMPORT FROM JSON
function import-calendar-items-from-json {
    param(
        [parameter(mandatory)]
        $FOLDER,

        [parameter(mandatory)]
        [string]$INPUT_PATH
    )

    $ROWS = get-content -path $INPUT_PATH -raw | convertfrom-json

    foreach ($ROW in $ROWS) {
        $NEW_ITEM = $FOLDER.Items.Add(1)

        $NEW_ITEM.Subject = $ROW.subject
        $NEW_ITEM.Start = [datetime]$ROW.start
        $NEW_ITEM.End = [datetime]$ROW.end
        $NEW_ITEM.Location = $ROW.location
        $NEW_ITEM.Body = $ROW.body
        $NEW_ITEM.Categories = $ROW.categories
        $NEW_ITEM.BusyStatus = [int]$ROW.busy_status
        $NEW_ITEM.ReminderSet = [bool]$ROW.reminder_set
        $NEW_ITEM.ReminderMinutesBeforeStart = [int]$ROW.reminder_minutes
        $NEW_ITEM.AllDayEvent = [bool]$ROW.alldayevent
        $NEW_ITEM.Sensitivity = [int]$ROW.sensitivity
        $NEW_ITEM.Importance = [int]$ROW.importance

        $NEW_ITEM.Save()
        write-host ("Created | {0} | {1}" -f ([datetime]$NEW_ITEM.Start), $NEW_ITEM.Subject)
    }
}

## 🦄 MAIN LOGIC: MENU
function start-calendar-tool {
    $NAMESPACE = get-outlook-namespace
    $FOLDER = select-calendar-folder -NAMESPACE $NAMESPACE

    while ($true) {
        write-host ""
        write-host "╔══════════════════════════════════════════════════════════════╗" -foregroundcolor darkcyan
        write-host "║                      CALENDAR CONTROL                        ║" -foregroundcolor cyan
        write-host "╚══════════════════════════════════════════════════════════════╝" -foregroundcolor darkcyan
        write-host ("📂 Selected: {0}" -f $FOLDER.FolderPath) -foregroundcolor cyan
        write-host ""
        write-host "[1] 📦 Export future items (JSON + CSV)" -foregroundcolor green
        write-host "[2] 👀 Preview one month" -foregroundcolor yellow
        write-host "[3] ❌ Delete one month" -foregroundcolor red
        write-host "[4] 👀 Preview one year" -foregroundcolor yellow
        write-host "[5] ❌ Delete one year" -foregroundcolor red
        write-host "[6] 📥 Import from JSON" -foregroundcolor magenta
        write-host "[7] 🔁 Change calendar" -foregroundcolor blue
        write-host "[Q] 🚪 Quit" -foregroundcolor gray
        write-host ""

        $MODE = read-host "Select command"

        if ($MODE -match '^(q|quit)$') {
            write-host "Goodbye." -foregroundcolor cyan
            return
        }

        if ($MODE -eq '7') {
            $NEW_FOLDER = select-calendar-folder -NAMESPACE $NAMESPACE -ALLOW_CANCEL
            if ($null -ne $NEW_FOLDER) {
                $FOLDER = $NEW_FOLDER
            }
            continue
        }

        try {
            switch ($MODE) {
                '1' {
                    $OUTPUT_PATH = read-host "Output path (file or folder)"
                    export-future-calendar-items -FOLDER $FOLDER -OUTPUT_PATH $OUTPUT_PATH
                }

                '2' {
                    $YEAR = read-int-in-range -PROMPT "Year" -MIN 1 -MAX 9999
                    $MONTH = read-int-in-range -PROMPT "Month number" -MIN 1 -MAX 12
                    $RANGE = get-month-range -YEAR $YEAR -MONTH $MONTH
                    show-calendar-items-in-range -FOLDER $FOLDER -START_DATE $RANGE.start_date -END_DATE $RANGE.end_date
                }

                '3' {
                    $YEAR = read-int-in-range -PROMPT "Year" -MIN 1 -MAX 9999
                    $MONTH = read-int-in-range -PROMPT "Month number" -MIN 1 -MAX 12
                    $RANGE = get-month-range -YEAR $YEAR -MONTH $MONTH
                    remove-calendar-items-in-range -FOLDER $FOLDER -START_DATE $RANGE.start_date -END_DATE $RANGE.end_date
                }

                '4' {
                    $YEAR = read-int-in-range -PROMPT "Year" -MIN 1 -MAX 9999
                    $RANGE = get-year-range -YEAR $YEAR
                    show-calendar-items-in-range -FOLDER $FOLDER -START_DATE $RANGE.start_date -END_DATE $RANGE.end_date
                }

                '5' {
                    $YEAR = read-int-in-range -PROMPT "Year" -MIN 1 -MAX 9999
                    $RANGE = get-year-range -YEAR $YEAR
                    remove-calendar-items-in-range -FOLDER $FOLDER -START_DATE $RANGE.start_date -END_DATE $RANGE.end_date
                }

                '6' {
                    $INPUT_PATH = read-host "Input JSON path"
                    import-calendar-items-from-json -FOLDER $FOLDER -INPUT_PATH $INPUT_PATH
                }

                default {
                    write-warning "Invalid option."
                }
            }
        }
        catch {
            write-warning $_.Exception.Message
        }
    }
}

## 🦄 PROCEDURE DIVISION: RUN
start-calendar-tool
