require 'axlsx'
Axlsx::Package.new do |p|
    p.workbook.add_worksheet(:name => "Pie Chart") do |sheet|
        sheet.add_row ["Simple Pie Chart"]
        %w(first second third).each { |label| sheet.add_row [label, rand(24)+1] }
        
        sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A6", :end_at => "F20", :grouping => :stacked,  :barDir => :col) do |chart|
            chart.add_series :data => sheet["B2:B4"], :labels => sheet["A2:A4"], :colors => ['FF0000', '00FF00', '0000FF']
        end
    end
    p.serialize('simple.xlsx')
end
