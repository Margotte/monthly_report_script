require 'rubyXL'
require 'rubyXL/convenience_methods'
require_relative 'support_features.rb'
require_relative 'monthly_report.rb'

# REMINDER SCRIPT

puts "Have you removed the extra IDs in the first column of your spreadsheet? (y|N)"
answer = STDIN.gets.chomp

return unless answer == "y"

# ---------------------------------------------------
# Create date
# ---------------------------------------------------

month = ARGV[0].to_i || 1 # default month is Jan
year = 2020
DATE = Date.new(year, month, 1)
Y_M_DATE = DATE.strftime("%Y-%m")
MONTHS_FR = [nil, "janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]


# ---------------------------------------------------
# Info is centralized in one excel document
# ---------------------------------------------------

origin_wkbk = RubyXL::Parser.parse("../../#{year}_paiements.xlsx")
origin_wksh = origin_wkbk.worksheets[0]


# ---------------------------------------------------
# Create the monthly excel document (coach workbook)
# ---------------------------------------------------

coach_wkbk = RubyXL::Workbook.new


# ---------------------------------------------------
# Create a new directory for month's reports
# ---------------------------------------------------
# ::mkdir returns 0

unless Dir.exist?("../#{Y_M_DATE}")
  Dir.mkdir("../#{Y_M_DATE}")
end

File.delete('#{Y_M_DATE}/#{Y_M_DATE}_coaches.xlsx') if File.exists? '#{Y_M_DATE}/#{Y_M_DATE}_coaches.xlsx'


# ---------------------------------------------------
# List of columns to be copied
# ---------------------------------------------------

ORIGIN_COLUMNS = list_columns(origin_wksh)

COLUMN_TITLES = ["date", "activité", "durée activité", "durée préparation", "durée totale",  "compensation brute"]
COLUMN_INDEXES = COLUMN_TITLES.map { |title| ORIGIN_COLUMNS.index(title) }

COACH_INDEX = ORIGIN_COLUMNS.index("animateur")
DATE_INDEX = ORIGIN_COLUMNS.index("date")
REGIME_INDEX = ORIGIN_COLUMNS.index("régime")

CURRENCY_COLUMNS = ["compensation brute"]

#  indexes in monthly report (not original ss)
CURRENCY_INDEXES = CURRENCY_COLUMNS.map { |title| COLUMN_TITLES.index(title) }

# ---------------------------------------------------
# Create list of coaches active on given month
# ---------------------------------------------------

coaches = list_coaches(origin_wksh)


# ---------------------------------------------------
# Create a worksheet for each coach
# ---------------------------------------------------

create_coaches_worksheets(coaches, coach_wkbk)


# ---------------------------------------------------
# Fill each coach's worksheet
# ---------------------------------------------------

coaches.each do |coach_name|
  dest_wksh = coach_wkbk["#{Y_M_DATE}_#{coach_name}_relevé"]
  create_coach_report(coach_name, dest_wksh, origin_wksh)
end


# ---------------------------------------------------
# Create monthly summary
# ---------------------------------------------------

create_summary(origin_wksh,coach_wkbk, coaches)

# rename Sheet1
summary_wksh = coach_wkbk.worksheets[0]
summary_wksh.sheet_name = "#{Y_M_DATE}_summary"


# ---------------------------------------------------
# Save to new workbook
# ---------------------------------------------------

coach_wkbk.write("../#{Y_M_DATE}/#{Y_M_DATE}_coaches.xlsx")
