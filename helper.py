from datetime import datetime, timedelta


# TR 8:10am-9:30am --> ['8:10am','9:30am']
# str -> list of str
def extract_meeting_times(input):
    # Split the time slot string based on space and hyphen
    parts = input.split(' ')
    if parts[0] == "Does":
        return ["NaN","NaN"]
    return parts[1].split('-')
    # Extract day and time parts


# TR 8:10am-9:30am --> TR
# str -> str
def extract_meeting_day(input):
    # Split the time slot string based on space and hyphen
    parts = input.split(' ')

    if parts[0] == "Does":
        return "NaN"
    # Extract day and time parts
    return parts[0]
    #day = parts[0]

# Perks, Gary (000000464) [100%] --> Perks, Gary
# str -> str
def extract_instructor(input):
    inst = input.split()
    if inst[0] == "Staff":
        return "Staff"
    else:
        return inst[0] + ' ' + inst[1]
    
# Frank E. Pilling 014 -0303 --> 014-0303
# str -> str
def extract_room(input):
    room_check = input.split()
    if room_check[-2].isdigit():
        return room_check[-2] + room_check[-1]
    else:
        return 'University Lecture Room'


# Takes a list of dictionaries and returns the unique values associated with 'Instructor'
# List of Dictionaries -> List of Strings
def get_unique_instructors(data):
    # Create an empty set to store unique Instructor values
    unique_instructors = set()

    # Loop over the list of dictionaries
    for entry in data:
        # Check if the 'Instructor' key exists in the dictionary
        if 'Instructor' in entry:
            # Add the 'Instructor' value to the set
            # Call extract_instructor() --> Valid Instructor
            unique_instructors.add(extract_instructor(entry['Instructor']))

    # Convert the set to a list and return
    lst = list(unique_instructors)

    # Set 'Staff' to the last column
    if 'Staff' in lst:
        lst.remove('Staff')
        lst.sort()
        lst.append('Staff')
    else:
        lst.sort()

    return lst

# Takes a list of dictionaries and returns the unique values associated with 'Room'
# List of Dictionaries -> List of Strings
def get_unique_rooms(data):
    # Create an empty set to store unique room values
    unique_rooms = set()

    # Loop over the list of dictionaries
    for entry in data:
        # Check if the 'Room' key exists in the dictionary
        if 'Room' in entry:
            unique_rooms.add(extract_room(entry['Room']))

    # Convert the set to a list and return
    lst = list(unique_rooms)

    # Set 'University Lecture Room' to the last column
    if 'University Lecture Room' in lst:
        lst.remove('University Lecture Room')
        lst.sort()
        lst.append('University Lecture Room')
    else:
        lst.sort()

    return lst


# '9:40am' --> '09:30 AM'
# '1pm' --> '01:00 PM'
def round_class_time_helper(time_str):
    if ':' in time_str:
        time_obj = datetime.strptime(time_str, '%I:%M%p')
    else:
        time_obj = datetime.strptime(time_str, '%I%p')

    # Get the minutes part
    minutes = time_obj.minute

    # Determine how many minutes to add to round to the nearest 30-minute mark
    if minutes < 15:
        rounded_minutes = 0
    elif minutes < 45:
        rounded_minutes = 30
    else:
        rounded_minutes = 0
        time_obj += timedelta(hours=1)

    # Create a new time object with the rounded minutes
    rounded_time_obj = time_obj.replace(minute=rounded_minutes, second=0, microsecond=0)

    # Format the rounded time as a string in the desired format
    rounded_time_str = rounded_time_obj.strftime('%I:%M %p')

    return rounded_time_str

# 'TR 9:40am-11am' --> '09:30 AM', '11:00 AM'
def round_class_time(meeting_str):
    if '-' in meeting_str: # maybe unnec
        time_range_str = meeting_str.split(' ')[1] # ['TR', '9:40am-11am']
        start_time_str, end_time_str = time_range_str.split('-')

        return round_class_time_helper(start_time_str), round_class_time_helper(end_time_str)

    else: 
        return "Not Listed", "Not Listed"

# '09:30 AM', '10:30 AM' --> 2
# Returns number of cells the lab occupies
def find_number_of_cells(start_time_str, end_time_str):
    # Parse the time strings into datetime objects
    start_time = datetime.strptime(start_time_str, '%I:%M %p')
    end_time = datetime.strptime(end_time_str, '%I:%M %p')

    # Calculate the difference in minutes
    time_difference = (end_time - start_time).total_seconds() / 60

    return int(time_difference) % 30