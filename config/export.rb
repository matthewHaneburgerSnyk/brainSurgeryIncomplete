require 'creek'
require 'axlsx'
require 'axlsx_rails'
require 'rubygems'
require 'fileutils'


class Export

def initialize(file_name)
    @file_name = file_name
   
    
    def makeExport(verbatim_data, graph_data, segment_data, mindset_columns, mindset_types, mindset_titles)
        p = Axlsx::Package.new
        wb = p.workbook
       
        title = wb.styles.add_style(:font_name => "Calibri",
                :sz=> 13,
                :border=>Axlsx::STYLE_THIN_BORDER,
                :alignment=>{:horizontal => :center},
                :b=> true )
                                    
        body = wb.styles.add_style(:font_name => "Calibri",
                :sz=> 11,
                :border=>Axlsx::STYLE_THIN_BORDER,
                :alignment=>{:horizontal => :left, :wrap_text => true},
                :bg_color => "ffffff")
                                                               
                                                               
                                                            
                                                                                               
        header = wb.styles.add_style(:font_name => "Calibri",
                :sz=> 11,
                :border=>Axlsx::STYLE_THIN_BORDER,
                :alignment=>{:horizontal => :center, :wrap_text => true},
                :bg_color => "C0C0C0",
                :fg_color => "000000",
                :b=> true)
                
  header_title = wb.styles.add_style(:font_name => "Calibri",
                :sz=> 13,
                :border=>Axlsx::STYLE_THIN_BORDER,
                :alignment=>{:horizontal => :center, :wrap_text => true},
                :b=> true)
        
        
        #verbatim - 9 per row
        #graph - 7 per row
        #segments - 1 per row
        #Mindset columns - 1
        #types - 1
        #titles - 1
       
        worksheet_name = @file_name.truncate(30)
        wb.add_worksheet(:name=> "#{@file_name.truncate(30)}") do |sheet|
            sheet.add_row [@file_name], :style=>title
            sheet.add_row []
            sheet.add_row []
            
            sheet.add_row ["Segments"] ,:style=>header_title
            sheet.add_row ["Segment 1","Segment 2","Segment 3"], :style=>header
            sheet.add_row segment_data, :style=>body
            
            sheet.add_row []
            sheet.add_row []
            
            sheet.add_row ["Mindsets"],:style=>header_title
            sheet.add_row ["Mindset Title","Type","Column"] ,:style=>header
            
            
            mindset_count = mindset_titles.size
            i = 0
            mindset_titles.each do |mindset|
                
                sheet.add_row [mindset_titles[i], mindset_types[i], mindset_columns[i]]
                i += 1
            end
            
            sheet.add_row []
            sheet.add_row []
           sheet.add_row ["Verbatim Data"], :style=>header_title
           sheet.add_row ["Topic Code","Type","Survey Column", "Topic Title", "Topic Frame of Reference", "# of Ranking evaluations", "# of Ranking Statements Shown", "Mindset Y/N"], :style=>header
           verbatim_data.each do |row|
            sheet.add_row row, :style=>body
           
           end
           
           sheet.add_row []
           
           sheet.add_row ["Graph Data"], :style=>header_title
           sheet.add_row ["Graph Code","Graph Type","Topics", "Graph Title"] ,:style=>header
           graph_data.each do |row|
               sheet.add_row row , :style=>body
           end
           
           
           
            
            
            
            
        end # end worksheet do
   
    p.serialize(@file_name)
    file = @file_name.to_s
    FileUtils.mv @file_name, "./public/uploads/exports/#{@file_name.truncate(30)}_inputs.xlsx"
    
    
    end # End make export
    
    
    
    
    
    
    
end



end
