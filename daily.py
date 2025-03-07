import helper, compiled, remake

# Makes the 'MWF Labs' Sheet in 'output.xls'
def make_mwf_labs(sheet, unique_rooms, lab_data):
    compiled.add_time_column(sheet)
    compiled.add_columns(sheet, unique_rooms)

    for lab_class in lab_data:
        meeting_days = helper.extract_meeting_day(lab_class['Meeting Pattern'])
        if meeting_days in "MWF":
            start_time_rounded, end_time_rounded = helper.round_class_time(lab_class['Meeting Pattern'])
            instructor = helper.extract_instructor(lab_class['Instructor'])
            course = lab_class['Course']
            section = lab_class['Section #']
            #num_cells = helper.find_number_of_cells(start_time_rounded, end_time_rounded)
            num_cells = compiled.calculate_chunks(start_time_rounded, end_time_rounded)
            lab_room = helper.extract_room(lab_class['Room'])

            compiled.schedule_lab(sheet, unique_rooms, lab_room, start_time_rounded, end_time_rounded, instructor, course, section, num_cells, meeting_days)

    compiled.adjust_column_widths(sheet)
    # '''CELL FORMAT:
    # Instructotr
    # Course-Section #
    # 'lab'
    # '''

# Makes the 'TR Labs' Sheet in 'output.xls'
def make_tr_labs(sheet, unique_rooms, lab_data):
    compiled.add_time_column(sheet)
    compiled.add_columns(sheet, unique_rooms)

    for lab_class in lab_data:
        meeting_days = helper.extract_meeting_day(lab_class['Meeting Pattern'])
        if meeting_days in "TR":
            start_time_rounded, end_time_rounded = helper.round_class_time(lab_class['Meeting Pattern'])
            instructor = helper.extract_instructor(lab_class['Instructor'])
            course = lab_class['Course']
            section = lab_class['Section #']
            #num_cells = helper.find_number_of_cells(start_time_rounded, end_time_rounded)
            num_cells = compiled.calculate_chunks(start_time_rounded, end_time_rounded)
            lab_room = helper.extract_room(lab_class['Room'])

            compiled.schedule_lab(sheet, unique_rooms, lab_room, start_time_rounded, end_time_rounded, instructor, course, section, num_cells, meeting_days)


    compiled.adjust_column_widths(sheet)
    '''CELL FORMAT:
    Instructotr
    Course-Section #
    'lab'
    '''

# Makes the 'Instructor' Sheet in 'output.xls'
def make_instructors(sheet, unique_instructors, lec_lab_sem):
    compiled.add_time_column(sheet)
    compiled.add_columns_instructors(sheet, unique_instructors)

    for lab_class in lec_lab_sem:
        start_time_rounded, end_time_rounded = helper.round_class_time(lab_class['Meeting Pattern'])
        instructor = helper.extract_instructor(lab_class['Instructor'])
        course = lab_class['Course']
        section = lab_class['Section #']
        #num_cells = helper.find_number_of_cells(start_time_rounded, end_time_rounded)
        num_cells = compiled.calculate_chunks(start_time_rounded, end_time_rounded)
        lab_room = helper.extract_room(lab_class['Room'])
        component = lab_class['Component']
        day = helper.extract_meeting_day(lab_class['Meeting Pattern'])
        remake.remake_schedule_instructor(sheet, unique_instructors, lab_room, start_time_rounded, end_time_rounded, instructor, course, section, num_cells, component, day)

    compiled.adjust_column_widths(sheet)
    '''CELL FORMAT:
    Instructotr
    Course-Section #
    Component
    Meeting Pattern
    '''
    ...

# Makes the 'Lab View' Sheet in 'output.xls'
def make_lab_view(sheet, lab_data):
    compiled.add_columns_list_view(sheet)

    # Initialize the row index for writing to the sheet
    current_row_index = 2

    # Loop over the list of dictionaries and write data to the sheet
    for record in lab_data:
        # Skip entries where "Meeting Pattern" is "Does Not Meet"
        if record.get('Meeting Pattern') == "Does Not Meet":
            continue

        # Write ID value to the first column
        sheet.cell(row=current_row_index, column=1, value=record.get('Course'))

        # Write Section # value to the second column
        sheet.cell(row=current_row_index, column=2, value=record.get('Section #'))
        sheet.cell(row=current_row_index, column=3, value=helper.extract_instructor(record.get('Instructor')))
        sheet.cell(row=current_row_index, column=4, value=record.get('Component'))
        sheet.cell(row=current_row_index, column=5, value=helper.extract_room(record.get('Room')))

        sheet.cell(row=current_row_index, column=6, value=helper.extract_meeting_day(record.get('Meeting Pattern')))
        
        meeting_times = helper.extract_meeting_times(record.get('Meeting Pattern'))
        sheet.cell(row=current_row_index, column=7, value=meeting_times[0])
        sheet.cell(row=current_row_index, column=8, value=meeting_times[1])

        # Increment the row index only after writing valid data
        current_row_index += 1


    compiled.adjust_column_widths_LV(sheet)
    # CELL FORMAT:
    # Course, Section, Instructor, Room, Meeting Patter (Days, Start, End), other, other