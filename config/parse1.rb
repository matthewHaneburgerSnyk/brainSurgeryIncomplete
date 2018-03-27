
require 'axlsx'


p = Axlsx::Package.new
wb = p.workbook


wb.add_worksheet(:name => "Bar Chart") do |sheet|
    sheet.add_row ["A Simple Bar Chart"]
    %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
    sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A6", :end_at => "F20") do |chart|
        chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"], :title => sheet["A1"], :colors => ["00FF00", "0000FF"],  :series_type => "PercentPositive.crtx"
    end
    p.serialize("example.xlsx")
end
