def list_columns(origin_wksh)
  origin_wksh.sheet_data.rows[0].cells.reject {|cell| cell.nil? || cell.value.nil?}.map { |cell| cell.value }
end


def add_title(dest_wksh, coach_name)
  dest_wksh.add_cell(0, 0, "#{DATE.strftime("%Y-%m")} TECHIES LAB")
  dest_wksh.add_cell(dest_wksh.count, 0, "Relevé mensuel des heures pour #{coach_name}")

  # To be able to later on style the two first rows, create the other cells of the rows:
  (1..COLUMN_TITLES.count-1).each do |column_index|
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
  currency_cell.set_number_format('[$€-80C] #.00')
end

def style_row(dest_wksh, row_number, styles = {})
  font = styles[:font] || 'Work Sans Light'
  font_size = styles[:font_size] || 12
  bold = styles[:bold] || false
  height = styles[:height] || 20
  horizontal_align = styles[:horizontal_align] || 'left'
  color = styles[:color]
  wrap = styles[:wrap]

  dest_wksh.change_row_font_size(row_number, font_size)
  dest_wksh.change_row_font_name(row_number, font)
  dest_wksh.change_row_bold(row_number, bold)
  dest_wksh.change_row_height(row_number, height)
  dest_wksh.change_row_horizontal_alignment(row_number, horizontal_align)

  if color
    (0..COLUMN_TITLES.count-1).each do |column_index|
      dest_wksh.sheet_data[row_number][column_index].change_fill(color)
    end
  end

  if wrap
    (0..COLUMN_TITLES.count-1).each do |column_index|
      dest_wksh.sheet_data[row_number][column_index].change_text_wrap(true)
    end
  end
end

def create_coach_report(coach_name, dest_wksh, origin_wksh)

  # Adding title of worksheet and blank row
  add_title(dest_wksh, coach_name)

  # Adding header row
  add_header_row(dest_wksh, origin_wksh)

  # Adding (copying) coaches' info
  add_info_rows(coach_name, dest_wksh, origin_wksh)

  # Adding total amount row
  add_total_row(dest_wksh)

  # style of each cells
  dest_wksh.count.times do |row|
    style_row(dest_wksh, row, {horizontal_align: 'center'})
  end

  # style of title (first two) rows
  style_row(dest_wksh, 0, {font: 'Space Mono', bold: true, height:  30, color: '5c8fc1'})
  style_row(dest_wksh, 1, {font: 'Space Mono', bold: true, color: '5c8fc1'})

  # style of header (third) row
  style_row(dest_wksh, 3, { bold: true, height:  35, wrap: true, horizontal_align: 'center'})

  # style of TOTAL row
  last_row = dest_wksh.count - 1
  style_row(dest_wksh, last_row, { bold: true, height:  30, color: 'b9cfe4', font: 'Space Mono', horizontal_align: 'center'})

  # width of columns
  {0 => 12, 1 => 18, 2 => 10, 3 => 12, 4 => 9, 5 => 11}.each do |column, width|
    dest_wksh.change_column_width(column, width)
  end
end
