require 'rubyXL'
require 'rubyXL/convenience_methods'
require_relative 'support_features.rb'
require_relative 'monthly_report.rb'

# DEBUG SCRIPT

# File.delete('2020-02/2020-01_coaches.xlsx') if File.exists? '2020-02/2020-01_coaches.xlsx'

# ---------------------------------------------------
# Info is centralized in one excel document
# ---------------------------------------------------

summary_wkbk = RubyXL::Parser.parse("../../2020_paiements.xlsx")
summary_wksh = summary_wkbk.worksheets[0]


# ---------------------------------------------------
# Create the monthly excel document (coach workbook)
# ---------------------------------------------------

coach_wkbk = RubyXL::Workbook.new


# ---------------------------------------------------
# Create date
# ---------------------------------------------------

month = ARGV[0].to_i || 1 # default month is Jan
year = 2020
DATE = Date.new(year, month, 1)


# ---------------------------------------------------
# Create a new directory for month's reports
# ---------------------------------------------------
# ::mkdir returns 0

Dir.mkdir("../#{DATE.strftime("%Y-%m")}")

# ---------------------------------------------------
# List of columns to be copied
# ---------------------------------------------------

ORIGIN_COLUMNS = list_columns(summary_wksh)
COLUMN_TITLES = ["date", "activité", "durée activité", "durée préparation", "durée totale", "total des frais"]
COLUMN_INDEXES = COLUMN_TITLES.map { |title| ORIGIN_COLUMNS.index(title) }
COACH_INDEX = ORIGIN_COLUMNS.index("animateur")
DATE_INDEX = ORIGIN_COLUMNS.index("date")

# ---------------------------------------------------
# Create list of coaches active on given month
# ---------------------------------------------------

coaches = list_coaches(summary_wksh)


# ---------------------------------------------------
# Create a worksheet for each coach
# ---------------------------------------------------

create_coaches_worksheets(coaches, coach_wkbk)


# ---------------------------------------------------
# Fill each coach's worksheet
# ---------------------------------------------------

coaches.each do |coach_name|
  dest_wksh = coach_wkbk["#{DATE.strftime("%Y-%m")}_#{coach_name}"]
  create_coach_report(coach_name, dest_wksh, summary_wksh)
end

# ---------------------------------------------------
# Save to new workbook
# ---------------------------------------------------

# remove Sheet1
coach_wkbk.worksheets.delete_at(0)

coach_wkbk.write("../2020-02/#{DATE.strftime("%Y-%m")}_coaches.xlsx")

