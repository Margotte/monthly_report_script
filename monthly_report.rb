def list_columns(origin_wksh)
  origin_wksh.sheet_data.rows[0].cells.reject {|cell| cell.nil? || cell.value.nil?}.map { |cell| cell.value }
end

def add_title(dest_wksh, title, column_count)
  dest_wksh.add_cell(0, 0, "#{DATE.strftime("%Y-%m")} TECHIES LAB")
  dest_wksh.add_cell(dest_wksh.count, 0, title)

  # To be able to later on style the two first rows, create the other cells of the rows:
  (1..column_count-1).each do |column_index|
    dest_wksh.add_cell(0, column_index, "")
    dest_wksh.add_cell(1, column_index, "")
  end

  # Blank row:
  dest_wksh.add_row(dest_wksh.count)
end

def add_header_row(dest_wksh, origin_wksh)
  next_row = dest_wksh.count
  header_row_values = COLUMN_INDEXES.map { |index| origin_wksh.sheet_data.rows.first.cells[index].value }

  header_row_values.each_with_index do |value, index|
    dest_wksh.add_cell(next_row, index, value)
  end
end

def add_info_rows(coach_name, dest_wksh, origin_wksh)

  origin_wksh.sheet_data.rows.each do |row|

    if row[COACH_INDEX].value == coach_name && row[DATE_INDEX].value.month == DATE.month

      next_row = dest_wksh.count

      row_values = COLUMN_INDEXES.map { |index| row.cells[index].value }

      row_values.each_with_index do |value, index|
        dest_wksh.add_cell(next_row, index, value)
      end

      # change format of date cells
      date_cell = dest_wksh.sheet_data.rows[dest_wksh.count - 1].cells.first
      date_cell.set_number_format('yyyy/mm/dd')

      # change format of currency cells
      currency_cell = dest_wksh.sheet_data.rows[dest_wksh.count - 1].cells.last
      currency_cell.set_number_format('[$€-80C] ##.00')
    end
  end
end

def add_total_row(dest_wksh)
  dest_wksh.add_cell(dest_wksh.count, 0, 'TOTAL')
  (1..3).each do |column_index|
    dest_wksh.add_cell((dest_wksh.count - 1), column_index, "")
  end
  dest_wksh.add_cell((dest_wksh.count - 1), 4, '', "SUM(E2:E#{dest_wksh.count - 1 })")
  dest_wksh.add_cell((dest_wksh.count - 1), 5, '', "SUM(F2:F#{dest_wksh.count - 1 })")
  currency_cell = dest_wksh.sheet_data.rows[dest_wksh.count - 1].cells.last
  currency_cell.datatype = RubyXL::DataType::NUMBER
  currency_cell.set_number_format('[$€-80C] #.00')
end

def style_row(dest_wksh, row_number, column_count, styles = {})
  font = styles[:font] || 'Work Sans Light'
  font_size = styles[:font_size] || 12
  bold = styles[:bold] || false
  height = styles[:height] || 20
  horizontal_align = styles[:horizontal_align] || 'left'
  color = styles[:color]
  wrap = styles[:wrap]
  border = styles[:border]

  dest_wksh.change_row_font_size(row_number, font_size)
  dest_wksh.change_row_font_name(row_number, font)
  dest_wksh.change_row_bold(row_number, bold)
  dest_wksh.change_row_height(row_number, height)
  dest_wksh.change_row_horizontal_alignment(row_number, horizontal_align)

  if color
    (0..column_count-1).each do |column_index|
      dest_wksh.sheet_data[row_number][column_index].change_fill(color)
    end
  end

  if wrap
    (0..column_count-1).each do |column_index|
      dest_wksh.sheet_data[row_number][column_index].change_text_wrap(true)
    end
  end

  if border
    (0..column_count-1).each do |column_index|
      dest_wksh.sheet_data[row_number][column_index].change_border(:top, 'thin')
      dest_wksh.sheet_data[row_number][column_index].change_border_color(:top, 'b9cfe4')
    end
  end
end

def style_cells(worksheet, column_count)
  worksheet.count.times do |row|
    style_row(worksheet, row, column_count, {horizontal_align: 'center'})
  end
end

def style_title_rows(worksheet, column_count)
  style_row(worksheet, 0, column_count, {font: 'Space Mono', bold: true, height:  30, color: '5c8fc1'})
  style_row(worksheet, 1, column_count, {font: 'Space Mono', bold: true, color: '5c8fc1'})
end

def width_columns(worksheet, widths)
  widths.each do |column, width|
    worksheet.change_column_width(column, width)
  end
end

def create_coach_report(coach_name, dest_wksh, origin_wksh)

  # Adding title of worksheet and blank row
  title = "Relevé mensuel des heures pour #{coach_name}"
  column_count = COLUMN_TITLES.count
  add_title(dest_wksh, title, column_count)

  # Adding header row
  add_header_row(dest_wksh, origin_wksh)

  # Adding (copying) coaches' info
  add_info_rows(coach_name, dest_wksh, origin_wksh)

  # Adding total amount row
  add_total_row(dest_wksh)

  # style of each cells
  style_cells(dest_wksh, column_count)

  # style of title (first two) rows
  style_title_rows(dest_wksh, column_count)

  # style of header (third) row
  style_row(dest_wksh, 3, column_count, { bold: true, height:  35, wrap: true, horizontal_align: 'center'})

  # style data rows
  (4..dest_wksh.count-2).each do |row|
    style_row(dest_wksh, row, column_count, {border: true, horizontal_align: 'center'})
  end

  # style of TOTAL row
  last_row = dest_wksh.count - 1
  style_row(dest_wksh, last_row, column_count, { bold: true, height:  30, color: 'b9cfe4', font: 'Space Mono', horizontal_align: 'center'})

  # width of columns
  widths = {0 => 12, 1 => 18, 2 => 10, 3 => 12, 4 => 9, 5 => 11}
  width_columns(dest_wksh, widths)
end

def create_summary(coach_wkbk, coaches)
  summary_wksh = coach_wkbk.worksheets[0]
  column_count = 4

  # Adding header row
  next_row = summary_wksh.count
  summary_wksh.add_cell(next_row, 0, "name")
  summary_wksh.add_cell(next_row, 1, "hours")
  summary_wksh.add_cell(next_row, 2, "fees")
  summary_wksh.add_cell(next_row, 3, "month")

  # Add info
  coaches.each do |coach_name|
    coach_wksh = coach_wkbk["#{DATE.strftime("%Y-%m")}_#{coach_name}_relevé"]
    summary_next_row = summary_wksh.count
    coach_last_row = coach_wksh.count - 2

    # name
    summary_wksh.add_cell(summary_next_row, 0, coach_name)

    # hours (UGLY FIX)
    hours_column = COLUMN_TITLES.index("durée totale")
    hours = (4..coach_last_row).map {|row_number| coach_wksh[row_number][hours_column].value }.sum

    summary_wksh.add_cell(summary_next_row, 1, hours)

    # fees (UGLY FIX)
    fees_column = COLUMN_TITLES.index("total des frais")
    fees = (4..coach_last_row).map {|row_number| coach_wksh[row_number][fees_column].value }.sum

    summary_wksh.add_cell(summary_next_row, 2, fees)
    currency_cell = summary_wksh[summary_next_row][2]
    currency_cell.set_number_format('[$€-80C] #.00')

    # month
    summary_wksh.add_cell(summary_next_row, 3, "#{MONTHS_FR[DATE.month]} #{DATE.year}")
  end

  style_cells(summary_wksh, column_count)
  style_row(summary_wksh, 0, column_count, { bold: true, height:  35, wrap: true, horizontal_align: 'center', color: 'b9cfe4'})

  # width of columns
  widths = {0 => 20, 1 => 15, 2 => 15, 3 => 20}
  width_columns(summary_wksh, widths)
end
