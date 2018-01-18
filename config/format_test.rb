require 'creek'
require 'axlsx'
require 'axlsx_rails'
require 'rubygems'
wb.add_worksheet(:name => "Auto Filter") do |sheet|
  sheet.add_row ["Build Matrix"]
  sheet.add_row ["Build", "Duration", "Finished", "Rvm"]
  sheet.add_row ["19.1", "1 min 32 sec", "about 10 hours ago", "1.8.7"]
  sheet.add_row ["19.2", "1 min 28 sec", "about 10 hours ago", "1.9.2"]
  sheet.add_row ["19.3", "1 min 35 sec", "about 10 hours ago", "1.9.3"]
  sheet.auto_filter = "A2:D5"
end
