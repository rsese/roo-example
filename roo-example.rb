# Author: Robert Sese
# Date: 10/25/2011
#
# Description: Simple example using the Roo gem to read an Excel spreadsheet
# and print out the data to standard out.  You would of course do something
# more interesting with the data depending on your application.

require 'roo'

workbook = Excel.new('a-list-apart-web-design-survey-sample.xls')

# Set the worksheet you want to work with as the default worksheet.  You could
# also iterate over all the worksheets in the workbook.
workbook.default_sheet = workbook.sheets[0]

# Create a hash of the headers so we can access columns by name (assuming row
# 1 contains the column headings).  This will also grab any data in hidden
# columns.
headers = Hash.new
workbook.row(1).each_with_index {|header,i|
  headers[header] = i
}

# Iterate over the rows using the `first_row` and `last_row` methods.  Skip
# the header row in the range.
((workbook.first_row + 1)..workbook.last_row).each do |row|

  # Get the column data using the column heading.
  age = workbook.row(row)[headers['What is your age in years?']]
  gender = workbook.row(row)[headers['What is your gender?']]
  most_identify_with = workbook.row(row)[headers['With which of these groups do you most identify?']]
  global_region = workbook.row(row)[headers['In which global region are you located?']]
  country = workbook.row(row)[headers['In which country are you located?']]
  education = workbook.row(row)[headers['What is the highest level of education you have completed?']]
  academics_helpfulness = workbook.row(row)[headers['How much have your academic studies helped you in your web work?']]

  print "Row: #{age}, #{gender}, #{most_identify_with}, #{global_region}, #{country}, #{education}, #{academics_helpfulness}\n\n"

end
