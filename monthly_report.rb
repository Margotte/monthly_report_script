def list_columns(origin_wksh)
  origin_wksh.sheet_data.rows[0].cells.reject {|cell| cell.nil? || cell.value.nil?}.map { |cell| cell.value }
end


def add_title(dest_wksh, coach_name)
  dest_wksh.add_cell(0, 0, "#{DATE.strftime("%Y-%b")}_#{coach_name}")
  (1..COLUMN_TITLES.count-1).each do |column_index|
    dest_wksh.add_cell(0, column_index, "")
  end
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
    dest_wksh.change_row_font_size(row, 12)
    dest_wksh.change_row_height(row, 20)
    dest_wksh.change_row_font_name(row, 'Work Sans Light')
    dest_wksh.change_row_horizontal_alignment(row, 'center')
  end

  # style of title (first) row
  dest_wksh.change_row_font_name(0, 'Space Mono')
  dest_wksh.change_row_bold(0, true)
  dest_wksh.change_row_height(0, 30)
  dest_wksh.change_row_horizontal_alignment(0, 'left')

  (0..COLUMN_TITLES.count-1).each do |column_index|
    dest_wksh.sheet_data[0][column_index].change_fill('5c8fc1')
  end

  # style of header (third) row
  dest_wksh.change_row_height(2, 35)
  dest_wksh.change_row_bold(2, true)
  (0..COLUMN_TITLES.count-1).each do |column_index|
    dest_wksh.sheet_data[2][column_index].change_text_wrap(true)
  end

  # style of TOTAL row
  last_row = dest_wksh.count - 1
  dest_wksh.change_row_font_name(last_row, 'Space Mono')
  dest_wksh.change_row_bold(last_row,true)
  dest_wksh.change_row_height(last_row, 30)
  dest_wksh.change_row_height(last_row, 30)
  (0..COLUMN_TITLES.count-1).each do |column_index|
    dest_wksh.sheet_data[last_row][column_index].change_fill('b9cfe4')
  end
  # dest_wksh.change_row_border(last_row, :top, 'medium')

  # style of columns
  {0 => 12, 1 => 18, 2 => 12, 3 => 15, 4 => 10, 5 => 12}.each do |column, width|
    dest_wksh.change_column_width(column, width)
  end
end


