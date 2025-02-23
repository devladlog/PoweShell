# Save this script as ComplimentScript.ps1

# Get the current hour
$currentHour = (Get-Date).Hour

# Determine the phase of the day and give a compliment
if ($currentHour -ge 5 -and $currentHour -lt 12) {
    $compliment = "Good morning! Just wanted to say, you're going to have an amazing day!"
} elseif ($currentHour -ge 12 -and $currentHour -lt 17) {
    $compliment = "Good afternoon! You're doing great, keep it up!"
} elseif ($currentHour -ge 17 -and $currentHour -lt 21) {
    $compliment = "Good evening! You've accomplished so much today, well done!"
} else {
    $compliment = "Good night! Rest well, you deserve it!"
}

# Add the SpeechSynthesizer class
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Set the voice to Microsoft Zira
$synthesizer.SelectVoice("Microsoft Zira Desktop")

# Speak the compliment
$synthesizer.Speak($compliment)

# Add the SpeechRecognitionEngine class
Add-Type -AssemblyName System.Speech
$speechRecognizer = New-Object System.Speech.Recognition.SpeechRecognitionEngine
$speechRecognizer.SetInputToDefaultAudioDevice()

# Define the grammar for recognizing commands
$grammarBuilder = New-Object System.Speech.Recognition.GrammarBuilder
$grammarBuilder.Append("create task")
$grammarBuilder.Append("create meeting")
$grammarBuilder.Append("create note")
$grammar = New-Object System.Speech.Recognition.Grammar($grammarBuilder)
$speechRecognizer.LoadGrammar($grammar)

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Calendars.ReadWrite"

# Get the calendar ID
$calendar = Get-MgUserCalendar -UserId "vladimiro.luis@devladlog.com" | Where-Object { $_.Name -eq "Calendar" }
$calendarId = $calendar.Id

# Get today's events from the calendar
$today = Get-Date
$startOfDay = $today.Date.ToString("yyyy-MM-ddTHH:mm:ss")
$endOfDay = $today.Date.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")

$filter = "start/dateTime ge '$startOfDay' and start/dateTime lt '$endOfDay'"
$events = Get-MgUserCalendarEvent -UserId "vladimiro.luis@devladlog.com" -CalendarId $calendarId -Filter $filter

# Function to strip HTML tags from content
function Remove-HTMLTags {
    param ($htmlContent)
    return [regex]::Replace($htmlContent, '<[^>]+>', '')
}

# Filter events for the current hour
$currentEvents = $events | Where-Object {
    $eventStart = [datetimeoffset]::Parse($_.Start.DateTime)
    $eventStart.Hour -eq $currentHour
}

# Speak the current tasks and meetings
if ($currentEvents.Count -gt 0) {
    $synthesizer.Speak("Here are your tasks and meetings for this hour:")
    foreach ($event in $currentEvents) {
        if ($event.Subject) {
            $synthesizer.Speak($event.Subject)
        }
        if ($event.Body.Content) {
            $plainTextContent = Remove-HTMLTags -htmlContent $event.Body.Content
            $synthesizer.Speak("Message body: " + $plainTextContent)
        }
    }
} else {
    $synthesizer.Speak("You have no tasks or meetings scheduled for this hour.")
}

# Check for any meetings
$meetings = $events | Where-Object {
    $_.IsMeeting -eq $true
}

# Speak the meetings
if ($meetings.Count -gt 0) {
    $synthesizer.Speak("You have the following meetings today:")
    foreach ($meeting in $meetings) {
        if ($meeting.Subject) {
            $synthesizer.Speak($meeting.Subject)
        }
        if ($meeting.Body.Content) {
            $plainTextContent = Remove-HTMLTags -htmlContent $meeting.Body.Content
            $synthesizer.Speak("Message body: " + $plainTextContent)
        }
    }
} else {
    $synthesizer.Speak("You have no meetings scheduled for today.")
}

# Get the next day's tasks and meetings
$nextDay = $today.AddDays(1)
$startOfNextDay = $nextDay.Date.ToString("yyyy-MM-ddTHH:mm:ss")
$endOfNextDay = $nextDay.Date.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")

$nextDayFilter = "start/dateTime ge '$startOfNextDay' and start/dateTime lt '$endOfNextDay'"
$nextDayEvents = Get-MgUserCalendarEvent -UserId "vladimiro.luis@devladlog.com" -CalendarId $calendarId -Filter $nextDayFilter

# Speak the next day's tasks and meetings if it's night time (e.g., after 8 PM)
if ($currentHour -ge 20) {
    $synthesizer.Speak("Tasks and meetings for tomorrow:")
    foreach ($event in $nextDayEvents) {
        if ($event.Subject) {
            $synthesizer.Speak($event.Subject)
        }
        if ($event.Body.Content) {
            $plainTextContent = Remove-HTMLTags -htmlContent $event.Body.Content
            $synthesizer.Speak("Message body: " + $plainTextContent)
        }
    }

    # Provide advice to close the day
    $closingAdvice = @(
        "Take a moment to reflect on your achievements today.",
        "Prepare your to-do list for tomorrow.",
        "Relax and unwind with a good book or a favorite show.",
        "Get a good night's sleep to recharge for tomorrow."
    )
    $synthesizer.Speak("Advice to close the day: " + ($closingAdvice | Get-Random))
}

# Function to create a task
function Create-Task {
    $taskSubject = Read-Host "Enter the task subject"
    $taskStart = Read-Host "Enter the task start time (yyyy-MM-ddTHH:mm:ss)"
    $taskEnd = Read-Host "Enter the task end time (yyyy-MM-ddTHH:mm:ss)"
    New-MgUserCalendarEvent -UserId "vladimiro.luis@devladlog.com" -CalendarId $calendarId -Subject $taskSubject -Start @{DateTime=$taskStart; TimeZone="UTC"} -End @{DateTime=$taskEnd; TimeZone="UTC"}
    $synthesizer.Speak("Task created successfully.")
}

# Function to create a meeting
function Create-Meeting {
    $meetingSubject = Read-Host "Enter the meeting subject"
    $meetingStart = Read-Host "Enter the meeting start time (yyyy-MM-ddTHH:mm:ss)"
    $meetingEnd = Read-Host "Enter the meeting end time (yyyy-MM-ddTHH:mm:ss)"
    $attendees = Read-Host "Enter the attendees' email addresses (comma-separated)"
    $attendeeList = $attendees -split ","
    $attendeeObjects = @()
    foreach ($attendee in $attendeeList) {
        $attendeeObjects += @{EmailAddress=@{Address=$attendee}}
    }
    New-MgUserCalendarEvent -UserId "vladimiro.luis@devladlog.com" -CalendarId $calendarId -Subject $meetingSubject -Start @{DateTime=$meetingStart; TimeZone="UTC"} -End @{DateTime=$meetingEnd; TimeZone="UTC"} -Attendees $attendeeObjects
    $synthesizer.Speak("Meeting created successfully.")
}

# Function to create a note
function Create-Note {
    $noteContent = Read-Host "Enter the note content"
    # Save the note content to a file or a note-taking application
    $notePath = "C:\Scripts\Notes.txt"
    Add-Content -Path $notePath -Value $noteContent
    $synthesizer.Speak("Note created successfully.")
}

# Handle speech recognized events
$speechRecognizer.Add_SpeechRecognized({
    param ($sender, $e)
    switch ($e.Result.Text) {
        "Zira create task" { Create-Task }
        "Zira create meeting" { Create-Meeting }
        "Zira create note" { Create-Note }
    }
})

# Start asynchronous speech recognition
$speechRecognizer.RecognizeAsync([System.Speech.Recognition.RecognizeMode]::Multiple)