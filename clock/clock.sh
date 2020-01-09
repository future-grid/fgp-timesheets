#!/bin/bash

# if the directory doesn't exist, make it
source=$(realpath "${0}" | xargs dirname)
filepath=$HOME/.timesheets/
template=template.xlsx
timestore=.clock
folderfile=.parent
[[ ! -d $filepath ]] && mkdir -p $filepath && echo "Created $filepath"
[[ ! -f $filepath$template ]] && echo "Template not found!" && exit

if [ ! -f $filepath$folderfile ]; then
    echo "Probably your first run, saving necessary google drive stuff."
    id=$(gdrive list --query "sharedWithMe" | grep "WeeklyTimeSheets" | awk '{print $1}')
    [[ -z $id ]] && echo "Couldn't get timesheets folder!" && exit
    echo $id > $filepath$folderfile
else
    parent=$(cat $filepath$folderfile)
fi

projects=( "FGP-19.10" "FGP-PGI" "WEL-E" "WEL-SIQ" "WEL-SUP" "UI-UIU" "UE-FGPU" "UE-C" "UE-NP" "UE-SUP" "CPS-O" "CPS-SUP" "TA" "BSA" "SH" "ADM" "WP" "SAPN" "PL-3" "PL-4" "PL-5" "PL-6" )
titles=( "FGP - 19.10" "FGP - PG Image" "WEL - Event" "WEL - SensorIQ" "WEL - Support" "UE - UI Upgrade" "UE - platform upgrade" "UE - Consumption" "UE - Network Planning" "UE - Support" "CPS - OnSight" "CPS - Support" "ThingsAt" "BSA" "Shinehub" "Admin" "Western Power" "SAPN" "Place Holder" "Place Holder" "Place Holder" "Place Holder" )
### DO NOT EDIT ABOVE THIS POINT!!! stuff will break ###

if [ -z $projects ]; then
    python3 projects.py $filepath$template
    echo "Please re-run the script!"
    exit
fi
# personal details
initials=AP

# put the two together
friday="20"$(date -v+Friday '+%y%m%d')
timesheet="TimeSheet_w$friday"_"$initials".xlsx
now=$(date '+%s')

# arrow key support for menu
ReadCharacter() {
    major=""
    minor=""
#   essentially when you enter in an arrow key youre actually entering in a string of 3
    read -rsn1 ui
    case "$ui" in
        $'\x1b')    # Handle ESC sequence.
            read -rsn1  tmp 
            if [[ "$tmp" == "[" ]]; then
                read -rsn1  tmp 
                case "$tmp" in
                    "A") major="up";;
                    "B") major="down";;
                    "C") major="right";;
                    "D") major="left";;
                esac
            fi
            read -rsn5 -t 0
            ;;
        "") minor="enter";;
        [Jj]) minor="down";;
        [Kk]) minor="up";;
    esac

    if [[ $minor == "" ]]; then
        char="$major"
    else
        char="$minor"
    fi
}
movepos() {
    ((position+=$1))
}
ProjectMenu() {
    position=0
    while true; do
        clear
        echo "Select project"
        echo "--------------"

        len=${#projects[@]}
        ((lensub=len-1))

        # checks for values below 0, and above the maximum
        [[ $position -gt $lensub ]] && position=$lensub
        [[ $position -lt 0 ]] && position=0
        for index in ${!projects[@]}; do
            if [[ $index = $position ]]; then
                echo "${projects[$index]}: ${titles[$index]} <--"
            else
                echo "${projects[$index]}: ${titles[$index]}"
            fi
        done

        ReadCharacter
        case $char in
            "up") movepos -1 ;;
            "down") movepos 1 ;;
            "enter") project=${projects[$position]} && break ;;
        esac
    done
}
DescriptionWriter() {
    clear
    echo "A description of your time spent:"
    echo "---------------------------------"
    read description
    description=$(echo $description | sed 's/ /~/g')
}

# if there are two bits
if [ $# -gt 1 ]; then
    if [ $1 == "in" ]; then
        chosentime=$2
        min=$3
        sec=$4
    fi
fi


ClockIn() {
    if [ ! -z $chosentime ]; then
      now=$(date -j -f "%H:%M:%S" "$chosentime:${min:-00}:${sec:-00}" "+%s")
      time=$(date -j -f "%s" "$now" "+%H:%M")
      echo "Special time chosen: $time"
    fi
    # create the usable spreadsheet
    [[ ! -f $filepath$timesheet ]] && cp $filepath$template $filepath$timesheet
    if [ ! -f $filepath$timestore ]; then
        echo $now > $filepath$timestore
    else
        echo "Clock in time already exists! Overwrite? (y/n)"
        while true; do
            read -s -n1 reply
            case $reply in 
                y) echo $now > $filepath$timestore && echo "Overwritten." && break ;;
                Y) echo $now > $filepath$timestore && echo "Overwritten." && break ;;
                n) echo "Not overwriting, exiting" && break ;;
                N) echo "Not overwriting, exiting" && break ;;
            esac
        done
    fi

    # flavor
    time=$(date -j -f "%s" "$now" "+%H:%M")
    echo "Clocked in at $time"
    exit
}
# will be called normally. get the time theyve been here eg. 'check the clock'
GetTime() {
    [[ ! -f $filepath$timestore ]] && echo "You haven't clocked in yet!" && exit

    then=$(cat $filepath$timestore)
    difference=$(( $now - $then ))
    clocked_at=$(date -r "$then" +%H:%M) # convert timestamp clocked time to readable time

    # calculate the hours in 30 minute blocks
    hours=$(echo "scale=2; $difference/3600" | bc)
    hours=$(echo "(($hours+0.25)/0.5)*0.5" | bc)

    echo "Hours: $hours"
    echo "You clocked in at: $clocked_at"
    exit
}
UploadTimesheet() {
    month=$(date +%m-%B)
    year=$(date +%Y)
    year_folder=$(gdrive list --query " '$parent' in parents" | grep $year-weekly-timesheet | awk '{print $1}')
    month_folder=$(gdrive list --query " '$year_folder' in parents" | grep $month | awk '{print $1}')

    [[ -z $parent ]] && echo "Do you have gdrivecli installed?" && exit
    if [ -z $year_folder ]; then
        echo "Generating year folder."
        gdrive mkdir --parent $parent $year-weekly-timesheet 
        year_folder=$(gdrive list --query " '$parent' in parents" | grep $year-weekly-timesheet | awk '{print $1}')
    fi
    if [ -z $month_folder ]; then
        echo "Generating month folder."
        gdrive mkdir --parent $year_folder $month
        month_folder=$(gdrive list --query " '$year_folder' in parents" | grep $month | awk '{print $1}')
    fi
    gdrive upload --parent $month_folder $filepath$timesheet
    echo "Finished!"
}

ClockOut() {
    [[ ! -f $filepath$timestore ]] && echo "You need to clock in first!" && exit

    then=$(cat $filepath$timestore)
    difference=$(( $now - $then ))

    # calculate the hours in 30 minute blocks
    hours=$(echo "scale=2; $difference/3600" | bc)
    hours=$(echo "(($hours+0.25)/0.5)*0.5" | bc)
     
    ProjectMenu
    DescriptionWriter
    today=$(date '+%A')

    pyscript="clock.py"
    
    python3 $source/$pyscript $filepath$timesheet $hours $project $today $description
    rm $filepath$timestore

     if [ $today == "Friday" ]; then
         echo "It's friday! Would you like to Google Drive the timesheet?"
         while true; do
             read -s -n1 reply
             case $reply in 
                 y) UploadTimesheet;;
                 Y) UploadTimesheet;;
                 n) echo "Done" && exit;;
                 N) echo "Done" && exit;;
             esac
         done
     fi
}

# actual code execution part!

if [ $# -eq 0 ]; then
    GetTime
fi

for arg in "$@"
do
    case "$arg" in 
        in) ClockIn;;
        out) ClockOut;;
        upload) UploadTimesheet;;
    esac
done
