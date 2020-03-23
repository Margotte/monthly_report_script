# create an array of coaches

def list_coaches(worksheet)
  coaches = []

  (1...worksheet.count).each do |row|
    coach_cell = worksheet[row][COACH_INDEX]
    date_cell = worksheet[row][DATE_INDEX]
    if date_cell.value.month == DATE.month && !coach_cell.nil?
      coaches << coach_cell.value
    end
  end

  return coaches.uniq
end

# create an individual sheet for each coach

def create_coaches_worksheets(coaches, workbook)
  coaches.each do |coach_name|
    coach_worksheet = workbook.add_worksheet("#{DATE.strftime("%Y-%m")}_#{coach_name}_relevé")
  end
end



