require 'creek'
require 'axlsx'
require 'axlsx_rails'
require 'rubygems'
require 'fileutils'





class Parser
 
def initialize(mapping)
#Parse doc incoming
$file_name = mapping[1].to_s
$file_name = $file_name.gsub(" ", "_")
$file_name = $file_name.gsub("(", "")
$file_name = $file_name.gsub(")", "")
cworkbook = Creek::Book.new mapping[0]
cworksheets = cworkbook.sheets


puts "Found #{cworksheets.count} worksheets"

$data = Array.new




cworksheets.each do |cworksheet|
    puts "Reading: #{cworksheet.name}"
    num_rows = 0

    
    
   
    cworksheet.rows.each do |row|
        row_cells = row.values
        num_rows += 1
        #puts "Full row for debugging #{row_cells}"
        $data.push(row.values.join "          ")
        
        # uncomment to print out row values
        #puts row_cells.join "    "
        
    end
  

  
    data_col = Array.new
    col_one = Array.new
    
    


end
###############################################################################################################################################
######## make verbatim start ##################################################################################################################
###############################################################################################################################################
#def makeVerb(topic_code, topic_type, survey_column, topic_title, topic_frame_of_reference, segment, product_mindset , ranking_num, ranking_total)

def makeVerb( values , segments, mindsets, mindset_types, mindset_titles )
run_count = values.size
$run_count_table = run_count

p = Axlsx::Package.new
wb = p.workbook

values.each do |verbatim_call|
    


topic_code               = verbatim_call[0]
topic_type               = verbatim_call[1]
survey_column            = verbatim_call[2]
topic_title              = verbatim_call[3]
topic_frame_of_reference = verbatim_call[4]
ranking_num              = verbatim_call[5]
ranking_total            = verbatim_call[6]

distance_counter = 0
distance = 3

$main_distance_row = $data[0].to_s.split("          ")
start_spot = $main_distance_row.index(survey_column)
$g_start_spot = $main_distance_row.index(survey_column)
#segment_calc = distance_row.index(segment)
#product_start_spot = distance_row.index(product_mindset)
$rank_num = ranking_num
$rank_total = ranking_total



#remove empty segments
#There can be up to three segments
$segments        = segments.reject { |v| v.to_s.empty? }
$segment_count   = $segments.size
$segment_s_array  = Array.new

$segments.each do |seg|
   $segment_s_array.push($main_distance_row.index(seg))
end



#remove empty mindsets
#there can be up to five mindsets
puts
$mindsets        = mindsets.reject { |v| v.to_s.empty? }
$mindset_count   = $mindsets.size
$mindset_s_array = Array.new
$mindsets.each do |mindset|
    $mindset_s_array.push($main_distance_row.index(mindset))
end


puts "mindsets #{$mindsets}"


puts "mindset_s_array #{$mindset_s_array}"

$mindset_types = mindset_types.reject { | v| v.to_s == "Select Type" }
$mindset_types_array = Array.new
$mindset_types.each do |type|
    $mindset_types_array.push(type)
    
end

$mindset_titles = mindset_titles.reject { |v| v.to_s.empty? }
$mindset_titles_array = Array.new
$mindset_titles.each do |title|
    $mindset_titles_array.push(title)
    
end


segment_calc               = $segment_s_array[1]
product_start_spot         = $mindset_s_array[1]


#verbatim styles
title = wb.styles.add_style(:font_name => "Calibri",
                            :sz=> 16,
                            :border=>Axlsx::STYLE_THIN_BORDER,
                            :alignment=>{:horizontal => :center},
                            :b=> true )
                            
body = wb.styles.add_style(:font_name => "Calibri",
                           :sz=> 11,
                           :border=>Axlsx::STYLE_THIN_BORDER,
                           :alignment=>{:horizontal => :left, :wrap_text => true},
                           :bg_color => "ffffff")
                           
                           
                           
intensity = wb.styles.add_style(:font_name => "Calibri",
                           :sz=> 11,
                           :border=>Axlsx::STYLE_THIN_BORDER,
                           :alignment=>{:horizontal => :center, :wrap_text => true},
                           :bg_color => "ffffff")
                           
                           
header = wb.styles.add_style(:font_name => "Calibri",
                           :sz=> 11,
                           :border=>Axlsx::STYLE_THIN_BORDER,
                           :alignment=>{:horizontal => :center, :wrap_text => true},
                           :bg_color => "C0C0C0",
                           :fg_color => "000000",
                           :b=> true)


def valence_calc(value)
   case value.to_i
    when 1
      return "Very Pleasant"
    when 2
      return "Mildly Pleasant"
    when 3
      return "Neutral"
    when 4
      return "Mildly Unpleasant"
    when 5
      return "Very Unpleasant"
   end
end


def map_mindset_value(mindset_value)
    if mindset_value < 1.5
        mindset = "Impassioned"
        elsif mindset_value < 2.6
        mindset = "Attracted"
        elsif mindset_value < 3.5
        mindset = "Apathetic"
        else
        mindset = "Unattracted"
        
        return mindset
    end
end


def mindset_calc( valence1, valence2, valence3 )
   
   mindsetc = valence1.to_i + valence2.to_i + valence3.to_i
   mindsetc = mindsetc / 3
   
   #return mindsetc
  return map_mindset_value(mindsetc)
   
end



def r_calc( col_val, statement_rank, type  )
    #gets types below from their respective columns
    #col_val = value in the column, 2, 3, 4 or whatever, tells the function which valence/statement to get
    #statement_rank = number of statements, maps to ranking total
    #statement_eval = number of statements evaluated
    #type = emotion/intensity/why/valence
    
    #type calc vals
    type_val = 0
    case type
        when type = "emotion"
           type_val = 2
        when type = "intensity"
           type_val = 3
        when type = "why"
           type_val = 4
        when type = "valence"
           type_val = 5
        when type = "s"
           type_val = 5
    end
    #leave that 6 in there, it corresponds to how the mapping files are setup
    find_att   = (col_val.to_i * 6) - 6
    att_offset = statement_rank.to_i + $g_start_spot
    
    
    return find_att + att_offset + type_val

end


def r_mindset_calc( statement_rank, statement_count, row_num)
    data_row = $data[row_num].to_s.split("          ") # Parses individual rows from sheet
    #type is always valence
    #
    count = 0
    calc_total = 0
    while count < statement_rank.to_i
        
       r_calc_column = $g_start_spot + count
       
       r_calc_value = data_row[r_calc_column]
       
       
       if r_calc_value.to_i <= statement_count.to_i
           calc_total = data_row[r_calc(r_calc_value, statement_rank , "valence" )].to_i + calc_total
        
       end
       count = count + 1
    end
    
    mindsetc = calc_total / statement_count.to_i
    
    return map_mindset_value(mindsetc)
   
end



def many_mindset_calc(column, type, row_num)
    
    if column == nil || type == nil || row_num == nil
        
      return ""
    
    else
        
    data_row = $data[row_num].to_s.split("          ") # Parses individual rows from sheet
    
    
    #column is the index in the row, type is standard_3x or standard_1x
    if type == "standard_3x"
        p1 = column.to_i + 6
        p2 = column.to_i + 10
        p3 = column.to_i + 14
        
        mindset = data_row[p1].to_i + data_row[p2].to_i + data_row[p3].to_i
        mindset = mindset/3
        return map_mindset_value(mindset)
        
    
    elsif type == "standard_1x"
        p1 = column.to_i + 6
        mindset = data_row[p1]
        return map_mindset_value(mindset)
        
    else
    
      return ""
    
    end
    end
end


def segment_calc(segment, row_num)
   if segment.to_s.empty?
       return ""
   else
   data_row = $data[row_num].to_s.split("          ") # Parses individual rows from sheet
   
       return data_row[segment]
       
   end
    
end


# this is 3x only 
s11 = start_spot       # Emotion
s12 = start_spot + 4   # Intensity
s13 = start_spot + 5   # Why
s14 = start_spot + 6   # Valence
s21 = start_spot + 1   # Emotion
s22 = start_spot + 8   # Intensity
s23 = start_spot + 9   # Why
s24 = start_spot + 10  # Valence
s31 = start_spot + 2   # Emotion
s32 = start_spot + 12  # Intensity
s33 = start_spot + 13  # Why
s34 = start_spot + 14  # Valence



# this is 1x only
ss11 = start_spot       # Emotion
ss12 = start_spot + 1   # Intensity
ss13 = start_spot + 2   # Why
ss14 = start_spot + 3   # Valence

p14 = product_start_spot + 6   # Valence
p24 = product_start_spot + 10  # Valence
p34 = product_start_spot + 14  # Valence



row_count = 0
pop_row_count = 0



 worksheet_name = "#{topic_code} - #{topic_title}"
 worksheet_name = worksheet_name.truncate(30)
 worksheet_name = worksheet_name.gsub("/", "_")
 wb.add_worksheet(:name=> "#{worksheet_name}") do |sheet|
     
     #i_value = $data.size.to_i
     #i_value = i_value - 2
     #sheet.add_table "A3:I52"

#get 3x
case topic_type
  when "standard_3x"
   $data.each do |row|
        if row == $data[0]
            
        puts "Skipping Main Row"
        row_count = row_count + 1
        elsif row == $data[1]
        
        puts "Skipping second row"
        row_count = row_count + 1
        else
        data_row = $data[row_count].to_s.split("          ") # Parses individual rows from sheet
        
        #emotion, intensity, why, s, valence, mindset 1, segment 1, mindset 2, mindset 3, mindset 4, mindset 5, segment 2, segment 3, response
        
        
        
        #generates rows
        row_1_raw = [ data_row[s11] , #emotion
                      data_row[s12] ,  #Intensity
                      data_row[s13] ,  #Why
                      data_row[s14].to_i , #s
                      valence_calc(data_row[s14]) ,  #valence
                      mindset_calc(data_row[s14], data_row[s24], data_row[s34]), #mindset
                      many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ), #mindset 1
                      many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ), #mindset 2
                      many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ), #mindset 3
                      many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ), #mindset 4
                      many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ), #mindset 5
                      segment_calc($segment_s_array[0], row_count),   #segment 1
                      segment_calc($segment_s_array[1], row_count),   #segment 2
                      segment_calc($segment_s_array[2], row_count),   #segment 3
                      data_row[0]] # response id
        
        row_1 = row_1_raw.reject { |r| r.to_s.empty? }
        
        
        row_2_raw = [ data_row[s21] ,
                      data_row[s22] ,
                      data_row[s23] ,
                      data_row[s24].to_i ,
                      valence_calc(data_row[s24]),
                      mindset_calc(data_row[s14], data_row[s24], data_row[s34]), #mindset
                      many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ), #mindset 1
                      many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ), #mindset 2
                      many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ), #mindset 3
                      many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ), #mindset 4
                      many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ), #mindset 5
                      segment_calc($segment_s_array[0], row_count),   #segment 1
                      segment_calc($segment_s_array[1], row_count),   #segment 2
                      segment_calc($segment_s_array[2], row_count),   #segment 3
                      data_row[0]] # response id
                      
        row_2 = row_2_raw.reject { |r| r.to_s.empty? }
        
        row_3_raw =[ data_row[s31] ,
                     data_row[s32] ,
                     data_row[s33] ,
                     data_row[s34].to_i ,
                     valence_calc(data_row[s34]),
                     mindset_calc(data_row[s14], data_row[s24], data_row[s34]), #mindset
                     many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ), #mindset 1
                     many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ), #mindset 2
                     many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ), #mindset 3
                     many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ), #mindset 4
                     many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ), #mindset 5
                     segment_calc($segment_s_array[0], row_count),   #segment 1
                     segment_calc($segment_s_array[1], row_count),   #segment 2
                     segment_calc($segment_s_array[2], row_count),   #segment 3
                     data_row[0]] # response id
        
         row_3 = row_3_raw.reject { |r| r.to_s.empty? }
        
        
    
        sheet.add_row row_1, :style=>body
        
        sheet.add_row row_2, :style=>body
        
        sheet.add_row row_3, :style=>body
        row_count = row_count + 1
       end
        
        
       
    end
   pop_row_count = (row_count * 3) - 2
   
   sheet.rows.each do |row|
      
      row.cells[1].style = intensity
      
   end
   
   sheet.rows.sort_by!{ |row| [row.cells[3].value.to_i, -row.cells[1].value.to_i] }


#$mindset_s_array
#$segment_s_array
#$mindset_types_array
puts "Mindset Titles array #{$mindset_titles_array}"

title_row_raw = [ "Emotion", #emotion
                  "Intensity" ,  #Intensity
                  "Why" ,  #Why
                  "S" , #s
                  "Valence" ,  #valence
                  "Mindset" ,
                  if many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ) == "" then "" else $mindset_titles_array[0] end, #mindset 1
                  if many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ) == "" then "" else $mindset_titles_array[1] end, #mindset 2
                  if many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ) == "" then "" else $mindset_titles_array[2] end, #mindset 3
                  if many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ) == "" then "" else $mindset_titles_array[3] end, #mindset 4
                  if many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ) == "" then "" else $mindset_titles_array[4] end, #mindset 5
                  "Segment",
                  if segment_calc($segment_s_array[1], row_count) == "" then "" else "Segment 2" end,   #segment 2
                  if segment_calc($segment_s_array[2], row_count) == "" then "" else "Segment 3" end,   #segment 3
                  "Response ID"] # response id
                  
              title_row = title_row_raw.reject { |t| t.to_s.empty? }


   sheet.add_row title_row, :style=>header
   sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])

   sheet.add_row [topic_frame_of_reference], :style=>title
   sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])

   sheet.add_row [topic_title], :style=>title
   sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
   
   
   sheet.merge_cells "A1:I1"
   sheet.merge_cells "A2:I2"
   sheet.column_widths 25, 25, 80, 5, 25, 25, 25, 25, 25
   
   # standard_1x
   when "standard_1x"
   $data.each do |row|
       if row == $data[0]
           
           puts "Skipping Main Row"
           row_count = row_count + 1
           elsif row == $data[1]
           
           puts "Skipping second row"
           row_count = row_count + 1
           else
           data_row = $data[row_count].to_s.split("          ") # Parses individual rows from sheet
           # increments count
           
           row_1_raw = [ data_row[ss11] , #emotion
                         data_row[ss12] ,  #Intensity
                         data_row[ss13] ,  #Why
                         data_row[ss14].to_i , #s
                         valence_calc(data_row[ss14]) ,  #valence
                         map_mindset_value(data_row[ss14]),
                         many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ), #mindset 1
                         many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ), #mindset 2
                         many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ), #mindset 3
                         many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ), #mindset 4
                         many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ), #mindset 5
                         segment_calc($segment_s_array[0], row_count),   #segment 1
                         segment_calc($segment_s_array[1], row_count),   #segment 2
                         segment_calc($segment_s_array[2], row_count),   #segment 3
                         data_row[0]] # response id
           
           row_1 = row_1_raw.reject { |r| r.to_s.empty? }


           sheet.add_row row_1, :style=>body
           
           
           
           row_count = row_count + 1
       end
       
   end
   pop_row_count = (row_count * 3) - 2
   
   sheet.rows.each do |row|
       
       row.cells[1].style = intensity
       
   end
   
   sheet.rows.sort_by!{ |row| [row.cells[3].value.to_i, -row.cells[1].value.to_i] }
   
   #sheet.add_row ["Emotion", "Intensity", "Why", "S" , "Valence", "Mindset", "Segment", "Product Mindset" ,"Response ID"], :style=>header
   
   title_row_raw = [ "Emotion", #emotion
                     "Intensity" ,  #Intensity
                     "Why" ,  #Why
                     "S" , #s
                     "Valence" ,  #valence
                     "Mindset", #mindset
                     if many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ) == "" then "" else $mindset_titles_array[0] end, #mindset 1
                     if many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ) == "" then "" else $mindset_titles_array[1] end, #mindset 2
                     if many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ) == "" then "" else $mindset_titles_array[2] end, #mindset 3
                     if many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ) == "" then "" else $mindset_titles_array[3] end, #mindset 4
                     if many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ) == "" then "" else $mindset_titles_array[4] end, #mindset 5
                     "Segment",   #segment 1
                     if segment_calc($segment_s_array[1], row_count) == "" then "" else "Segment 2" end,   #segment 2
                     if segment_calc($segment_s_array[2], row_count) == "" then "" else "Segment 3" end,   #segment 3
                     "Response ID"] # response id
   
   title_row = title_row_raw.reject { |t| t.to_s.empty? }
   
   
   sheet.add_row title_row, :style=>header
   sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
   
   sheet.add_row [topic_frame_of_reference], :style=>title
   sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
   
   sheet.add_row [topic_title], :style=>title
   sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
   
   
   sheet.merge_cells "A1:I1"
   sheet.merge_cells "A2:I2"
   sheet.column_widths 25, 25, 80, 5, 25, 25, 25, 25, 25
   
   #Ranking
   when "ranking"
   original_start = s11
   
   while s11 < original_start + $rank_total.to_i
   
   if s11 == original_start
       
       $data.each do |row|
           data_row = $data[row_count].to_s.split("          ") # Parses individual rows from sheet
           
           if $rank_num.to_i < data_row[s11].to_i
               puts "Not evaluated statement, skipping"
               row_count = row_count + 1
               
               elsif row == $data[0]
               
               puts "Skipping Main Row"
               row_count = row_count + 1
               elsif row == $data[1]
               
               puts "Skipping second row"
               row_count = row_count + 1
               else
               
               
               
     
     
  
                
                            

                            
                            
               row_1_raw = [data_row[r_calc(data_row[s11], $rank_total , "emotion" )] ,
                            data_row[r_calc(data_row[s11], $rank_total , "intensity" )].to_i,
                            data_row[r_calc(data_row[s11], $rank_total , "why" ) ],
                            data_row[r_calc(data_row[s11], $rank_total , "s" ) ].to_i,
                            valence_calc(data_row[r_calc(data_row[s11], $rank_total , "valence" )]),
                            r_mindset_calc( $rank_total, $rank_num, row_count), #mindset
                            many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ), #mindset 2
                            many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ), #mindset 3
                            many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ), #mindset 4
                            many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ), #mindset 5
                            segment_calc($segment_s_array[0], row_count),   #segment 1
                            segment_calc($segment_s_array[1], row_count),   #segment 2
                            segment_calc($segment_s_array[2], row_count),   #segment 3
                            data_row[0]]
                            
               
               row_1 = row_1_raw.reject { |r| r.to_s.empty? }

               
               sheet.add_row row_1, :style=>body
               
               
               row_count = row_count + 1
           end
       end
       
       sheet.rows.each do |row|
           
           row.cells[1].style = intensity
           
       end
       
       sheet.rows.sort_by!{ |row| [row.cells[3].value.to_i, -row.cells[1].value.to_i] }
       
       title_row_raw = [ "Emotion", #emotion
       "Intensity" ,  #Intensity
       "Why" ,  #Why
       "S" , #s
       "Valence" ,  #valence,
       "Mindset", #mindset
       if many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ) == "" then "" else $mindset_titles_array[0] end, #mindset 1
       if many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ) == "" then "" else $mindset_titles_array[1] end, #mindset 2
       if many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ) == "" then "" else $mindset_titles_array[2] end, #mindset 3
       if many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ) == "" then "" else $mindset_titles_array[3] end, #mindset 4
       if many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ) == "" then "" else $mindset_titles_array[4] end, #mindset 5
       "Segment",          #segment 1
       if segment_calc($segment_s_array[1], row_count) == "" then "" else "Segment 2" end,   #segment 2
       if segment_calc($segment_s_array[2], row_count) == "" then "" else "Segment 3" end,   #segment 3
       "Response ID"] # response id
       
       title_row = title_row_raw.reject { |t| t.to_s.empty? }
       
       
       sheet.add_row title_row, :style=>header
       sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
       
       sheet.add_row [topic_frame_of_reference], :style=>title
       sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
       
       sheet.add_row [topic_title], :style=>title
       sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
       
       
       sheet.merge_cells "A1:I1"
       sheet.merge_cells "A2:I2"
       sheet.column_widths 25, 25, 80, 5, 25, 25, 25, 25, 25
       
       
   else
   
   row_count = 0
   worksheet_name = "#{topic_code} - #{topic_title}_#{s11}"
   worksheet_name = worksheet_name.truncate(30)
   worksheet_name = worksheet_name.gsub("/", "_")
   wb.add_worksheet(:name=> "#{worksheet_name}") do |sheet|
       
       #sheet.add_row [topic_title], :style=>title
       
       # sheet.add_row [topic_frame_of_reference], :style=>title
       #sheet.merge_cells "A1:I1"
       #sheet.merge_cells "A2:I2"
       #sheet.add_row ["Emotion", "Intensity", "Why", "S" , "Valence", "Mindset", "Segment", "Product Mindset" ,"Response ID"]
       
       
       $data.each do |row|
           data_row = $data[row_count].to_s.split("          ") # Parses individual rows from sheet
           
           if $rank_num.to_i < data_row[s11].to_i
               puts "Not evaluated statement, skipping"
               row_count = row_count + 1
               
               elsif row == $data[0]
               
               puts "Skipping Main Row"
               row_count = row_count + 1
               elsif row == $data[1]
               
               puts "Skipping second row"
               row_count = row_count + 1
               else
               
               
               
               row_1_raw = [data_row[r_calc(data_row[s11], $rank_total , "emotion" )] ,
                            data_row[r_calc(data_row[s11], $rank_total , "intensity" )].to_i,
                            data_row[r_calc(data_row[s11], $rank_total , "why" ) ],
                            data_row[r_calc(data_row[s11], $rank_total , "s" ) ].to_i,
                            valence_calc(data_row[r_calc(data_row[s11], $rank_total , "valence" )]),
                            r_mindset_calc( $rank_total, $rank_num, row_count),
                            many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ), #mindset 2
                            many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ), #mindset 3
                            many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ), #mindset 4
                            many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ), #mindset 5
                            segment_calc($segment_s_array[0], row_count),   #segment 1
                            segment_calc($segment_s_array[1], row_count),   #segment 2
                            segment_calc($segment_s_array[2], row_count),   #segment 3
                            data_row[0]]
               
               
               row_1 = row_1_raw.reject { |r| r.to_s.empty? }
               
               
               sheet.add_row row_1, :style=>body
               
               
               row_count = row_count + 1
           end
       end
       sheet.rows.each do |row|
           
           row.cells[1].style = intensity
           
       end
       
       
       sheet.rows.sort_by!{ |row| [row.cells[3].value.to_i, -row.cells[1].value.to_i] }
       
       title_row_raw = [ "Emotion", #emotion
       "Intensity" ,  #Intensity
       "Why" ,  #Why
       "S" , #s
       "Valence" ,  #valence,
       "Mindset",
       if many_mindset_calc($mindset_s_array[0],$mindset_types_array[0], row_count ) == "" then "" else $mindset_titles_array[0] end, #mindset 1
       if many_mindset_calc($mindset_s_array[1],$mindset_types_array[1], row_count ) == "" then "" else $mindset_titles_array[1] end, #mindset 2
       if many_mindset_calc($mindset_s_array[2],$mindset_types_array[2], row_count ) == "" then "" else $mindset_titles_array[2] end, #mindset 3
       if many_mindset_calc($mindset_s_array[3],$mindset_types_array[3], row_count ) == "" then "" else $mindset_titles_array[3] end, #mindset 4
       if many_mindset_calc($mindset_s_array[4],$mindset_types_array[4], row_count ) == "" then "" else $mindset_titles_array[4] end, #mindset 5
       "Segment",                                              #segment 1
       if segment_calc($segment_s_array[1], row_count) == "" then "" else "Segment 2" end,   #segment 2
       if segment_calc($segment_s_array[2], row_count) == "" then "" else "Segment 3" end,   #segment 3
       "Response ID"] # response id
       
       title_row = title_row_raw.reject { |t| t.to_s.empty? }
       
       
       sheet.add_row title_row, :style=>header
       sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
       
       sheet.add_row [topic_frame_of_reference], :style=>title
       sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
       
       sheet.add_row [topic_title], :style=>title
       sheet.rows.insert 0,  sheet.rows.delete(sheet.rows[sheet.rows.length-1])
       
       
       sheet.merge_cells "A1:I1"
       sheet.merge_cells "A2:I2"
       sheet.column_widths 25, 25, 80, 5, 25, 25, 25, 25, 25
   end

   

  end
   s11 = s11 + 1
   pop_row_count = (row_count * 3) - 2
 end
   
   
   
   
   else puts "Not a known topic code"




end # end topic type case





sheet.add_table "A3:I#{pop_row_count-1}"



 end # end worksheet creator
#test commit number 2


end # end while to catch all sheets

p.serialize($file_name)
file = $file_name.to_s
FileUtils.mv $file_name, "./public/uploads/#{$file_name}/verbatim_#{$file_name}"




end



###############################################################################################################################################
########## make graph start ###################################################################################################################
###############################################################################################################################################

#def makeGraph(g_topic_code, g_topic_type, graph_topics, g_topic_title, g_ranking_num, g_ranking_total )
def makeGraph(values,segments, mindsets, mindset_types, mindset_titles)
    
    $segment_a = Array.new
    ######### Array positions of values ########
     #   valence = 4
     #   mindset = 5
     #   segment = 6
     #   product_mindset = 7
    
    #Open recently created verbatim and parse it
    verbatim = File.open("./public/uploads/#{$file_name}/verbatim_#{$file_name}")
    
    
    cworkbook = Creek::Book.new verbatim
    cworksheets = cworkbook.sheets
    puts "Found #{cworksheets.count} verbatim worksheets"
    
    # Where everything lives for graphs
    $gdata = Hash.new
    $o_gdata = Hash.new
    $m_gdata = Hash.new
    $om_gdata = Hash.new
    $g_data_raw = Array.new
    
    
    
    
    
    
    
    ######################### Parse out multiple segments/mindsets and their titles ##########################
    
    
    #remove empty segments
    #There can be up to three segments
    $segments        = segments.reject { |v| v.to_s.empty? }
    $segment_count   = $segments.size
    $segment_s_array  = Array.new
    $main_distance_row = $data[0].to_s.split("          ")
    
    $segments.each do |seg|
        $segment_s_array.push($main_distance_row.index(seg))
    end
    
    
    
    #remove empty mindsets
    #there can be up to five mindsets
    puts
    $mindsets        = mindsets.reject { |v| v.to_s.empty? }
    $mindset_count   = $mindsets.size
    $mindset_s_array = Array.new
    $mindsets.each do |mindset|
        $mindset_s_array.push($main_distance_row.index(mindset))
    end
    
    
    puts "mindsets #{$mindsets}"
    
    
    puts "mindset_s_array #{$mindset_s_array}"
    
    $mindset_types = mindset_types.reject { | v| v.to_s == "Select Type" }
    $mindset_types_array = Array.new
    $mindset_types.each do |type|
        $mindset_types_array.push(type)
        
    end
    
    $mindset_titles = mindset_titles.reject { |v| v.to_s.empty? }
    $mindset_titles_array = Array.new
    $mindset_titles.each do |title|
        $mindset_titles_array.push(title)
        
    end
    
    
    ########################################################################################################
    
    
    
    
    
    #################################   Percent Positive Calcs   ###############################
    #Parse verbatim by sheet
    
    
    
    $gdata      = Hash.new
    $m_gdata    = Hash.new
    $o_gdata    = Hash.new
    $om_gdata   = Hash.new
    
    
    cworksheets.each do |cworksheet|
        
        
        #emotion, intensity, why, s, valence, mindset,  mindset 1, mindset 2, mindset 3, mindset 4, mindset 5, segment 1, segment 2, segment 3, response
        #   0         1       2   3     4        5           6        7           8           9          10        11         12        13
        
        
        puts "Reading verbatims for graph generation, sheet: #{cworksheet.name}"
        num_rows = 0
        
        
    
        #get the contents of the sheet and put them in an array
        cworksheet.rows.each do |row|
            row_cells = row.values
            num_rows += 1
            $g_data_raw.push(row.values.join "          ")
        end
        
  
      
  
        #Percent Positive start values
        row_count = 0
        positive = 0
        neutral = 0
        negative = 0
        unattracted = 0
        apathetic = 0
        attracted = 0
        impassioned = 0
        #parse array to get values for building graphs
        $g_data_raw.each do |row|
            
            if row == $g_data_raw[0]
                
                puts "Skipping Main Row"
                row_count = row_count + 1
                elsif row == $g_data_raw[1]
                
                puts "Skipping second row"
                row_count = row_count + 1
                elsif row == $g_data_raw[2]
                puts "Skipping third row"
                row_count = row_count + 1
                
                else
                
                #emotion, intensity, why, s, valence, mindset,  mindset 1, mindset 2, mindset 3, mindset 4, mindset 5, segment1, segment 2, segment 3, response
                
                
                
                
                data_row = $g_data_raw[row_count].to_s.split("          ") # Parses individual rows from sheet
                row_size = data_row[0].size
                $mindset_count
                $segment_count
                
                if data_row[4] == "Very Pleasant"
                    positive = positive + 1
                    elsif data_row[4] == "Mildly Pleasant"
                    positive = positive + 1
                    elsif data_row[4] == "Neutral"
                    neutral = neutral + 1
                    elsif data_row[4] == "Mildly Unpleasant"
                    negative = negative + 1
                    elsif data_row[4] == "Very Unpleasant"
                    negative = negative + 1
                end
                
                if data_row[5] == "Unattracted"
                    unattracted = unattracted + 1
                    elsif data_row[5] == "Apathetic"
                    apathetic = apathetic + 1
                    elsif data_row[5] == "Attracted"
                    attracted = attracted + 1
                    elsif data_row[5] == "Impassioned"
                    impassioned = impassioned + 1
                    
                end






                
                #add segments to array so you can manipulate it later
                #$segment_a.push(data_row[7])
                #figure out how many segments there are and where they are
                #add them to the segments array
                $seg_spot == 0
                if $segment_count == 1
                    $seg_spot = $mindset_count + 6
                    $segment_a.push(data_row[$seg_spot])
                
                
                elsif $segment_count == 2
                    $seg_spot = $mindset_count + 6
                    $segment_a.push(data_row[$seg_spot])
                
                    $seg_spot_1 = $mindset_count + 7
                    $segment_a.push(data_row[$seg_spot_1])
                    
                elsif $segment_count == 3
                $seg_spot = $mindset_count + 6
                $segment_a.push(data_row[$seg_spot])
                
                $seg_spot_1 = $mindset_count + 7
                $segment_a.push(data_row[$seg_spot_1])
                
                $seg_spot_3 = $mindset_count + 8
                $segment_a.push(data_row[$seg_spot_3])
                
                end
                
                row_count = row_count + 1
            end
            
            
        end # End for parsing calculation
        #percent Positive
        $total = positive + neutral + negative
        $per_positive = (positive.to_f / $total.to_f).round(2)
        $per_neutral  = (neutral.to_f  / $total.to_f).round(2)
        $per_negative = (negative.to_f / $total.to_f).round(2)
        #mindset
        $m_total = unattracted + apathetic + attracted + impassioned
        $per_unattracted = (unattracted.to_f / $m_total.to_f).round(2)
        $per_apathetic   = (apathetic.to_f  / $m_total.to_f).round(2)
        $per_attracted   = (attracted.to_f / $m_total.to_f).round(2)
        $per_impassioned = (impassioned.to_f / $m_total.to_f).round(2)
        
        puts "per unattracted #{$per_unattracted}"
        puts "per apathetic #{$per_apathetic}"
        puts "per attracted  #{$per_attracted}"
        puts "per impassioned  #{$per_impassioned}"
    
        puts "per positive #{$per_positive}"
        puts "per negative #{$per_negative}"
        puts "per neutral  #{$per_neutral}"
        
        
        
    
        #################################   Segment Calcs   ###############################
        #Run segment calcs, this is all dynamic and should work with any number of segments
        row_count         = 0
        $seg_array        = Array.new
        $uniq_seg_names   = Array.new
        $uniq_segs        = Array.new
        $m_uniq_seg_names = Array.new
        $m_uniq_segs      = Array.new
        
        
        #how many segments are they and what are the unique vals
        $num_of_segs      = $segment_a.uniq.length
        $uniq_seg_names   = $segment_a.uniq
        $m_num_of_segs    = $segment_a.uniq.length
        $m_uniq_seg_names = $segment_a.uniq
        
        #turn segs values into a nested array
        $uniq_seg_names.each do |seg|
            array_spot = $uniq_seg_names.index(seg)
            $uniq_segs[array_spot] = Array.new
        end
        
        $m_uniq_seg_names.each do |seg|
            array_spot = $m_uniq_seg_names.index(seg)
            $m_uniq_segs[array_spot] = Array.new
        end


      #parse array to get values for building segment graphs
      $g_data_raw.each do |row|
            
            
            if row == $g_data_raw[0]

                puts "Skipping Main Row"
                row_count = row_count + 1
                
                elsif row == $g_data_raw[1]
                puts "Skipping second row"
                row_count = row_count + 1
                
                elsif row == $g_data_raw[2]
                puts "Skipping third row"
                row_count = row_count + 1
                
                else
                data_row = $g_data_raw[row_count].to_s.split("          ") # Parses individual rows from sheet
                
             
                $uniq_seg_names.each do |seg| # see what segment is on the row we're looking at
                    array_spot = $uniq_seg_names.index(seg)
                    if seg == data_row[7]
                       $uniq_segs[array_spot].push(data_row[4]) # Put the valence into an array corresponding to the segment title
                    end
                
               end
                $m_uniq_seg_names.each do |seg| # see what segment is on the row we're looking at
                    array_spot = $m_uniq_seg_names.index(seg)
                    if seg == data_row[7]
                        $m_uniq_segs[array_spot].push(data_row[5]) # Put the mindset into an array corresponding to the segment title
                    end
                    
                end
                
                
                row_count = row_count + 1
            end
            
      end # End segment row parsing
      
         $m_seg_values = Hash.new
         $m_uniq_seg_names.each do |seg|
             
             puts "M SEG NAME    =========== #{seg}"
             
             array_spot = $m_uniq_seg_names.index(seg)
             unattracted = 0
             apathetic = 0
             attracted = 0
             impassioned = 0
             
             temp_array = Array.new
             temp_array = $m_uniq_segs[array_spot]
             #$uniq_segs[array_spot].each do |val|
             temp_array.each do |val|
                 if val == "Unattracted"
                     unattracted = unattracted + 1
                     elsif val == "Apathetic"
                     apathetic = apathetic + 1
                     elsif val == "Attracted"
                     attracted = attracted + 1
                     elsif val == "Impassioned"
                     impassioned = impassioned + 1
                     
                 end
                 
                 
             end
             
             $s_m_total = unattracted + apathetic + attracted + impassioned
             $s_per_unattracted = (unattracted.to_f / $s_m_total.to_f).round(2)
             $s_per_apathetic   = (apathetic.to_f  / $s_m_total.to_f).round(2)
             $s_per_attracted   = (attracted.to_f / $s_m_total.to_f).round(2)
             $s_per_impassioned = (impassioned.to_f / $s_m_total.to_f).round(2)
             
             $m_seg_values[seg] = [$s_per_unattracted, $s_per_apathetic, $s_per_attracted, $s_per_impassioned, $s_m_total]
             
         end
      
      puts "m seg values =++++++++++++++++++ #{$m_seg_values}"
      
      
      
      
      
            $seg_values   = Hash.new
            $uniq_seg_names.each do |seg|
               array_spot = $uniq_seg_names.index(seg)
               positive = 0
               neutral = 0
               negative = 0
               
                temp_array = Array.new
                temp_array = $uniq_segs[array_spot]
                #$uniq_segs[array_spot].each do |val|
                temp_array.each do |val|
                    if val == "Very Pleasant"
                        positive = positive + 1
                        elsif val == "Mildly Pleasant"
                        positive = positive + 1
                        elsif val == "Neutral"
                        neutral = neutral + 1
                        elsif val == "Mildly Unpleasant"
                        negative = negative + 1
                        elsif val == "Very Unpleasant"
                        negative = negative + 1
                    end
                
                
               end
                $s_total = positive + neutral + negative
                $s_total_used = $total - $s_total
                $s_per_positive = (positive.to_f / $s_total.to_f).round(2)
                $s_per_neutral  = (neutral.to_f  / $s_total.to_f).round(2)
                $s_per_negative = (negative.to_f / $s_total.to_f).round(2)

                
               
                $seg_values[seg] = [$s_per_positive, $s_per_neutral, $s_per_negative, $s_total , $s_total_used]
    
                
            end
            
            
            
            # end # End for segment ID calculation
        
        
        
        
        
        ### That's what seg values looks like it's a hash with the key being the segment name
        #$seg_values[unique_seg_names(seg)] = (per_positive, per_neutral, per_negative, total_used, total)
  
        $seg_values         = $seg_values.reject { |k,v| k.nil? }
        $m_seg_values       = $m_seg_values.reject { |k,v| k.nil? }
   
        
        $ts   = [ $per_positive, $per_neutral, $per_negative,  $total, $total]
        $m_ts = [ $per_unattracted, $per_apathetic, $per_attracted, $per_impassioned, $m_total]
        
        puts "global ts #{$ts}"
        puts "global m_ts #{$m_ts}"
        
        
        title = cworksheet.name[0..2].strip
        m_title = cworksheet.name[0..2].strip
        $gdata[title]      = [$ts]
        $o_gdata[title]    = [$ts]
        $m_gdata[m_title]  = [$m_ts]
        $om_gdata[m_title] = [$m_ts]
        
      
        ##TODO - Filter out nil values from
         $seg_values.each do |key, value|
             title = cworksheet.name[0..2].strip
             
             
             if $gdata[title].include?(value)
                 puts "#{title} Already exists in doc, skipping"
             else
                 $gdata[title].push(key, value)
             end
             
             if $o_gdata[title].include?(value)
                 puts "#{title} Already exists in doc, skipping"
             else
                 $o_gdata[title].push(key, value)
             end
             #$gdata[title].push(key, value)
             #$o_gdata[title].push(key, value)
            
             
        end
     
         $m_seg_values.each do |key, value|
             m_title = cworksheet.name[0..2].strip
             
             
             if $m_gdata[m_title].include?(value)
                 puts "#{m_title} Already exists in doc, skipping"
                 else
                 $m_gdata[m_title].push(key, value)
             end
             
             if $om_gdata[m_title].include?(value)
                 puts "#{m_title} Already exists in doc, skipping"
                 else
                 $om_gdata[m_title].push(key, value)
             end
             #$m_gdata[m_title].push(key, value)
             #$om_gdata[m_title].push(key, value)
            
            end
       
      
      end # End for each worksheet

    #create new xlsx doc
    p = Axlsx::Package.new
    wb = p.workbook
    chart_style = wb.styles.add_style(:font_name => "Calibri",
                                      :sz=> 11,
                                      :fg_color => "000000",
                                      :format_code => "0%",
                                      :b=> true)
                                      
    chart_style_1 = wb.styles.add_style(:font_name => "Calibri",
                                      :sz=> 11,
                                      :fg_color => "000000",
                                      :b=> true)
                                      
                                      
                                      
  chart_style_2 = wb.styles.add_style(:font_name => "Calibri",
                                      :sz=> 11,
                                      :fg_color => "ffffff",
                                      :format_code => "00%",
                                      :border =>{:color => "00000000", :style => :none},
                                    
                                      :b=> true)
                                      
    
    puts "Graph Call+_+__+_+_+_+_+_+_+_+_+_+_+_+__+_+    #{values}"
    values.each do |graph_call|
        
        g_topic_code          = graph_call[0]
        g_topic_type          = graph_call[1]
        graph_topics          = graph_call[2]
        g_topic_title         = graph_call[3]
        g_ranking_num         = graph_call[4]
        g_ranking_total       = graph_call[5]
   
     if g_topic_type == "Percent Positive"
        
        $graph_topics_a = Array.new
        $graph_topics_a = graph_topics.split(",")
        $graph_topics_s = Array.new
        $graph_topics_s = graph_topics.split(",")

row_count = 0
$ts_positive    = Array.new
$ts_positive.push("Positive")
$ts_neutral     = Array.new
$ts_neutral.push("Neutral")
$ts_negative    = Array.new
$ts_negative.push("Negative")
$ts_total      = Array.new
$ts_total.push("Total")
$ts_total_used  = Array.new
$ts_total_used.push("Message Index")




$graph_topics_a.each do |topic|
    
 

    ts_data = $o_gdata[topic.strip]
    
    ts_data_a = ts_data[0]


    $ts_positive.push(ts_data_a[0])
    $ts_neutral.push(ts_data_a[1])
    $ts_negative.push(ts_data_a[2])
    $ts_total_used.push(ts_data_a[3])
    $ts_total.push(ts_data_a[4])
  
end





g_worksheet_name = "#{g_topic_code} (TS) - #{g_topic_title}"
g_worksheet_name = g_worksheet_name.truncate(30)
g_worksheet_name = g_worksheet_name.gsub("/", "_")
#  TS
puts "TS Worksheet name percent positive #{g_worksheet_name}"
wb.add_worksheet(:name=> "#{g_worksheet_name}") do |sheet|
    
    sheet.add_row [g_topic_title], :sz => 12, :font_name=>"Calibri", :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    sheet.add_row [g_topic_type], :sz => 12, :font_name=>"Calibri", :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    

            sheet.add_row  $graph_topics_a.unshift(""), :style => chart_style
            sheet.add_row  $ts_positive , :style => chart_style #positive
            sheet.add_row  $ts_neutral, :style => chart_style #neutral
            sheet.add_row  $ts_negative, :style => chart_style #negative
            sheet.add_row  $ts_total, :style => chart_style_1 #total
            sheet.add_row  $ts_total_used, :style => chart_style_1 #message index
            
            num_of_topics = $graph_topics_a.size - 1
           

            case num_of_topics
                when 1
                   $t_cells   = "B3:B3"
                   $p_cells   = "B4:B4"
                   $neu_cells = "B5:B5"
                   $neg_cells = "B6:B6"
                when 2
                   $t_cells   = "B3:C3"
                   $p_cells   = "B4:C4"
                   $neu_cells = "B5:C5"
                   $neg_cells = "B6:C6"
                when 3
                   $t_cells   = "B3:D3"
                   $p_cells   = "B4:D4"
                   $neu_cells = "B5:D5"
                   $neg_cells = "B6:D6"
                when 4
                   $t_cells   = "B3:E3"
                   $p_cells   = "B4:E4"
                   $neu_cells = "B5:E5"
                   $neg_cells = "B6:E6"
                when 5
                   $t_cells   = "B3:F3"
                   $p_cells   = "B4:F4"
                   $neu_cells = "B5:F5"
                   $neg_cells = "B6:F6"
                when 6
                   $t_cells   = "B3:G3"
                   $p_cells   = "B4:G4"
                   $neu_cells = "B5:G5"
                   $neg_cells = "B6:G6"
                when 7
                   $t_cells   = "B3:H3"
                   $p_cells   = "B4:H4"
                   $neu_cells = "B5:H5"
                   $neg_cells = "B6:H6"
                when 8
                   $t_cells   = "B3:I3"
                   $p_cells   = "B4:I4"
                   $neu_cells = "B5:I5"
                   $neg_cells = "B6:I6"
                when 9
                   $t_cells   = "B3:J3"
                   $p_cells   = "B4:J4"
                   $neu_cells = "B5:J5"
                   $neg_cells = "B6:J6"
                when 10
                   $t_cells   = "B3:K3"
                   $p_cells   = "B4:K4"
                   $neu_cells = "B5:K5"
                   $neg_cells = "B6:K6"
                when 11
                   $t_cells   = "B3:L3"
                   $p_cells   = "B4:L4"
                   $neu_cells = "B5:L5"
                   $neg_cells = "B6:L6"
                when 12
                   $t_cells   = "B3:M3"
                   $p_cells   = "B4:M4"
                   $neu_cells = "B5:M5"
                   $neg_cells = "B6:M6"
                when 13
                   $t_cells   = "B3:N3"
                   $p_cells   = "B4:N4"
                   $neu_cells = "B5:N5"
                   $neg_cells = "B6:N6"
                when 14
                   $t_cells   = "B3:O3"
                   $p_cells   = "B4:O4"
                   $neu_cells = "B5:O5"
                   $neg_cells = "B6:O6"
                when 15
                   $t_cells   = "B3:P3"
                   $p_cells   = "B4:P4"
                   $neu_cells = "B5:P5"
                   $neg_cells = "B6:P6"
                   
                else
                puts "Out of bounds, double check number of permitted topics per graph."
            end
      
     
     sheet.add_chart(Axlsx::Bar3DChart, :start_at => "B10", :end_at => "O32",:title=>"   ", :show_legend => false, :barDir => :col, :grouping => :percentStacked, ) do |chart|
              chart.add_series :data => sheet[$p_cells], :colors => ['365e92', '365e92', '365e92'], :labels => sheet[$t_cells], :color => "FFFFFF"
              chart.add_series :data => sheet[$neu_cells], :fg_color => "ffffff" , :colors => ['a5a5a5', 'a5a5a5', 'a5a5a5']
              chart.add_series :data => sheet[$neg_cells], :fg_color => "ffffff" , :colors => ['be0712', 'be0712', 'be0712']
              chart.d_lbls.show_val = true
              chart.d_lbls.show_percent = true
              chart.valAxis.gridlines = true
              chart.catAxis.gridlines = false
              chart.valAxis.format_code = "Percentage"
           
             
        


         
      end





end




#determine how many times to iterate
$topic_count = 0
$graph_topics_s.each do |topic|

    ts_data = $gdata[topic.strip]

    

      if ts_data[0].class == Array
        ts_data.delete_at(0)
    
      else
        puts "Not deleting segment titles"
      end
    
    count = ts_data.size
    $num_of_segments = count / 2
    $topic_count =+ 1
 end



puts "Number of segments #{$num_of_segments}"
$segment_name = Array.new
$ts_data = Array.new
$g_spot = 0
$seg_count = 0

while $seg_count < $num_of_segments
    $graph_topics_s = Array.new
    $graph_topics_s = graph_topics.split(",")

   
    #clear out array of ts values
    $ts_positive.clear
    $ts_positive.push("Positive")
    $ts_neutral.clear
    $ts_neutral.push("Neutral")
    $ts_negative.clear
    $ts_negative.push("Negative")
    $ts_total_used.clear
    $ts_total_used.push("Message Index")
    $ts_total.clear
    $ts_total.push("Total")
    #get values for segment

    
    puts " full gdata before #{$gdata}"
    
    $graph_topics_s.each do |topic|
        puts "topic being looked at #{topic}"
        $ts_data = $gdata[topic.strip]
        puts "gdata for debuggin #{$gdata[topic.strip]}"
        puts " full gdata after #{$gdata}"
        puts "ts data #{$ts_data}"
        if $seg_count == 0
            $g_spot = 1
            
            else
            
           $g_spot = ($seg_count * 2) + 1
        
        end
        
        ts_data_a = $ts_data[$g_spot.to_i]
        puts "TS Data a for segments #{ts_data_a}"
  
        $ts_positive.push(ts_data_a[0])
        $ts_neutral.push(ts_data_a[1])
        $ts_negative.push(ts_data_a[2])
        $ts_total_used.push(ts_data_a[3])
        $ts_total.push(ts_data_a[4])
  
        $name_count = $seg_count * 2

        $segment_name.push($ts_data[$name_count.to_i])

        $segment_name = $segment_name.uniq
     

        puts "segment name array #{$segment_name}"
        puts "segment name count #{$seg_count}"
    end
    
    $seg_title = "#{g_topic_code} #{$segment_name[$seg_count]} - #{g_topic_title}"
    $seg_title = $seg_title.truncate(30)
    $seg_title = $seg_title.gsub("/", "_")
 
 

#  Per Segments

puts "Segment worksheet name percent positive #{$seg_title}"
wb.add_worksheet(:name=> "#{$seg_title}") do |sheet|
    
    sheet.add_row [g_topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    sheet.add_row [$seg_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    
    
    
    sheet.add_row  $graph_topics_s.unshift("")
    sheet.add_row  $ts_positive, :style => chart_style #positive
    sheet.add_row  $ts_neutral, :style => chart_style  #neutral
    sheet.add_row  $ts_negative, :style => chart_style  #negative
    sheet.add_row  $ts_total, :style => chart_style_1 #total
    sheet.add_row  $ts_total_used, :style => chart_style_1 #message index
    
    
    
   
    num_of_topics = $graph_topics_s.size - 1

    
    case num_of_topics
        when 1
        $p_cells   = "B4:B4"
        $neu_cells = "B5:B5"
        $neg_cells = "B6:B6"
        when 2
        $p_cells   = "B4:C4"
        $neu_cells = "B5:C5"
        $neg_cells = "B6:C6"
        when 3
        $p_cells   = "B4:D4"
        $neu_cells = "B5:D5"
        $neg_cells = "B6:D6"
        when 4
        $p_cells   = "B4:E4"
        $neu_cells = "B5:E5"
        $neg_cells = "B6:E6"
        when 5
        $p_cells   = "B4:F4"
        $neu_cells = "B5:F5"
        $neg_cells = "B6:F6"
        when 6
        $p_cells   = "B4:G4"
        $neu_cells = "B5:G5"
        $neg_cells = "B6:G6"
        when 7
        $p_cells   = "B4:H4"
        $neu_cells = "B5:H5"
        $neg_cells = "B6:H6"
        when 8
        $p_cells   = "B4:I4"
        $neu_cells = "B5:I5"
        $neg_cells = "B6:I6"
        when 9
        $p_cells   = "B4:J4"
        $neu_cells = "B5:J5"
        $neg_cells = "B6:J6"
        when 10
        $p_cells   = "B4:K4"
        $neu_cells = "B5:K5"
        $neg_cells = "B6:K6"
        when 11
        $p_cells   = "B4:L4"
        $neu_cells = "B5:L5"
        $neg_cells = "B6:L6"
        when 12
        $p_cells   = "B4:M4"
        $neu_cells = "B5:M5"
        $neg_cells = "B6:M6"
        when 13
        $p_cells   = "B4:N4"
        $neu_cells = "B5:N5"
        $neg_cells = "B6:N6"
        when 14
        $p_cells   = "B4:O4"
        $neu_cells = "B5:O5"
        $neg_cells = "B6:O6"
        when 15
        $p_cells   = "B4:P4"
        $neu_cells = "B5:P5"
        $neg_cells = "B6:P6"
        
        else
        puts "Out of bounds, double check number of permitted topics per graph."
    end
  
  

    
    sheet.add_chart(Axlsx::Bar3DChart, :start_at => "B10", :end_at => "O32", :title=>"   ", :color => "000000", :show_legend => false, :barDir => :col, :grouping => :percentStacked, ) do |chart|
        chart.add_series :data => sheet[$p_cells],:fg_color => "ffffff" , :colors => ['365e92', '365e92', '365e92'], :labels => sheet[$t_cells]
        chart.add_series :data => sheet[$neu_cells], :fg_color => "ffffff" , :colors => ['a5a5a5', 'a5a5a5', 'a5a5a5']
        chart.add_series :data => sheet[$neg_cells], :fg_color => "ffffff" , :colors => ['be0712', 'be0712', 'be0712']
        chart.d_lbls.show_val = true
        chart.d_lbls.show_percent = true
        chart.valAxis.gridlines = true
        chart.catAxis.gridlines = false
        chart.valAxis.format_code = "Percentage"

$seg_count += 1

    end
    
end


end

end
   
end


puts "Values  +_+__+_+_+_+_+_+_+_+_+_+_+_+__+_+    #{values}"
values.each do |graph_call|
    puts "Graph Call+_+__+_+_+_+_+_+_+_+_+_+_+_+__+_+    #{graph_call}"
    gm_topic_code          = graph_call[0]
    gm_topic_type          = graph_call[1]
    m_graph_topics         = graph_call[2]
    gm_topic_title         = graph_call[3]
    gm_ranking_num         = graph_call[4]
    gm_ranking_total       = graph_call[5]
    
    if gm_topic_type == "Mindset"
        
        $m_graph_topics_a = Array.new
        $m_graph_topics_a = m_graph_topics.split(",")
        $m_graph_topics_s = Array.new
        $m_graph_topics_s = m_graph_topics.split(",")
        
        row_count = 0
        $ts_unattracted   = Array.new
        $ts_unattracted.push("Unattracted")
        $ts_apathetic     = Array.new
        $ts_apathetic.push("Apathetic")
        $ts_attracted    = Array.new
        $ts_attracted.push("Attracted")
        $ts_impassioned    = Array.new
        $ts_impassioned.push("Impassioned")
        $ts_total      = Array.new
        $ts_total.push("Total")
        
        puts "m_graph_topics_a ----------- #{$m_graph_topics_a}"
        puts "omg data #####   {#$om_gdata}"
        $m_graph_topics_a.each do |topic|
        
            ts_data = $om_gdata[topic.strip]
            ts_data_a = ts_data[0]
            puts "$$$$$$$$$$$$$$$$$$$$$  #{ts_data} $$$$$$$$$$$$$$$$$$$$$$$$"
            
            $ts_unattracted.push(ts_data_a[0])
            $ts_apathetic.push(ts_data_a[1])
            $ts_attracted.push(ts_data_a[2])
            $ts_impassioned.push(ts_data_a[3])
            $ts_total.push(ts_data_a[4])
            
       
       puts "Debugging all the mindset graph data"
       puts "unattracted #{$ts_unattracted}"
       puts "apathetic #{$ts_apathetic}"
       puts "attracted #{$ts_attracted}"
       puts "impassioned #{$ts_impassioned}"
       puts "total #{$ts_total}"
       
       
       
        end
        
        
        g_worksheet_name = "#{gm_topic_code} (TS) - #{gm_topic_title}"
        g_worksheet_name = g_worksheet_name.truncate(30)
        g_worksheet_name = g_worksheet_name.gsub("/", "_")
        #  TS
        
        puts "Mindset TS Worksheet name #{g_worksheet_name}"
        wb.add_worksheet(:name=> "#{g_worksheet_name}") do |sheet|
            
            sheet.add_row [gm_topic_title], :sz => 12, :font_name=>"Calibri", :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
            
            sheet.add_row [gm_topic_type], :sz => 12, :font_name=>"Calibri", :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
            
            
            #make title value
            if $m_graph_topics_a[0] == ""
            $m_graph_title = $m_graph_topics_a
            else
            $m_graph_title = $m_graph_topics_a.unshift("")
            end
            
            sheet.add_row  $m_graph_title, :style => chart_style
            sheet.add_row  $ts_unattracted , :style => chart_style
            sheet.add_row  $ts_apathetic, :style => chart_style
            sheet.add_row  $ts_attracted, :style => chart_style
            sheet.add_row  $ts_impassioned, :style => chart_style
            sheet.add_row  $ts_total, :style => chart_style_1
            
        
            num_of_topics = $m_graph_title.size - 1
            
            case num_of_topics
                when 1
                $t_cells   = "B3:B3"
                $un_cells  = "B4:B4"
                $ap_cells  = "B5:B5"
                $att_cells = "B6:B6"
                $imp_cells = "B7:B7"
                when 2
                $t_cells   = "B3:C3"
                $un_cells  = "B4:C4"
                $ap_cells  = "B5:C5"
                $att_cells = "B6:C6"
                $imp_cells = "B7:C7"
                when 3
                $t_cells   = "B3:D3"
                $un_cells  = "B4:D4"
                $ap_cells  = "B5:D5"
                $att_cells = "B6:D6"
                $imp_cells = "B7:D7"
                when 4
                $t_cells   = "B3:E3"
                $un_cells  = "B4:E4"
                $ap_cells  = "B5:E5"
                $att_cells = "B6:E6"
                $imp_cells = "B7:E7"
                when 5
                $t_cells   = "B3:F3"
                $un_cells  = "B4:F4"
                $ap_cells  = "B5:F5"
                $att_cells = "B6:F6"
                $imp_cells = "B7:F7"
                when 6
                $t_cells   = "B3:G3"
                $un_cells  = "B4:G4"
                $ap_cells  = "B5:G5"
                $att_cells = "B6:G6"
                $imp_cells = "B7:G7"
                when 7
                $t_cells   = "B3:H3"
                $un_cells  = "B4:H4"
                $ap_cells  = "B5:H5"
                $att_cells = "B6:H6"
                $imp_cells = "B7:H7"
                when 8
                $t_cells   = "B3:I3"
                $un_cells  = "B4:I4"
                $ap_cells  = "B5:I5"
                $att_cells = "B6:I6"
                $imp_cells = "B7:I7"
                when 9
                $t_cells   = "B3:J3"
                $un_cells  = "B4:J4"
                $ap_cells  = "B5:J5"
                $att_cells = "B6:J6"
                $imp_cells = "B7:J7"
                when 10
                $t_cells   = "B3:K3"
                $un_cells  = "B4:K4"
                $ap_cells  = "B5:K5"
                $att_cells = "B6:K6"
                $imp_cells = "B7:K7"
                when 11
                $t_cells   = "B3:L3"
                $un_cells  = "B4:L4"
                $ap_cells  = "B5:L5"
                $att_cells = "B6:L6"
                $imp_cells = "B7:L7"
                when 12
                $t_cells   = "B3:M3"
                $un_cells  = "B4:M4"
                $ap_cells  = "B5:M5"
                $att_cells = "B6:M6"
                $imp_cells = "B7:M7"
                when 13
                $t_cells   = "B3:N3"
                $un_cells  = "B4:N4"
                $ap_cells  = "B5:N5"
                $att_cells = "B6:N6"
                $imp_cells = "B7:N7"
                when 14
                $t_cells   = "B3:O3"
                $un_cells  = "B4:O4"
                $ap_cells  = "B5:O5"
                $att_cells = "B6:O6"
                $imp_cells = "B7:O7"
                when 15
                $t_cells   = "B3:P3"
                $un_cells  = "B4:P4"
                $ap_cells  = "B5:P5"
                $att_cells = "B6:P6"
                $imp_cells = "B7:P7"
                
                else
                puts "Out of bounds, double check number of permitted topics per graph."
            end
            
            
            sheet.add_chart(Axlsx::Bar3DChart, :start_at => "B10", :end_at => "O32",:title=>"   ", :show_legend => false, :barDir => :bar, :grouping => :percentStacked, ) do |chart|
                chart.add_series :data => sheet[$un_cells], :colors => ['be0712', 'be0712', 'be0712'], :labels => sheet[$t_cells], :color => "FFFFFF"
                chart.add_series :data => sheet[$ap_cells], :fg_color => "ffffff" , :colors => ['a5a5a5', 'a5a5a5', 'a5a5a5']
                chart.add_series :data => sheet[$att_cells], :fg_color => "ffffff" , :colors => ['ffffff', 'ffffff', 'ffffff']
                chart.add_series :data => sheet[$imp_cells], :fg_color => "ffffff" , :colors => ['365e92', '365e92', '365e92']
                chart.d_lbls.show_val = true
                chart.d_lbls.show_percent = true
                chart.valAxis.gridlines = true
                chart.catAxis.gridlines = false
                chart.valAxis.format_code = "Percentage"
                
                
                
                
                
            end
            
            
            
            
        end
        
        
        
        
        
        #determine how many times to iterate
        $topic_count = 0
        $m_graph_topics_s .each do |topic|
            ts_data = $m_gdata[topic.strip]
            if ts_data[0].class == Array
                ts_data.delete_at(0)
                else
                puts "Not deleting segment titles"
            end
            count = ts_data.size
            $num_of_segments = count / 2
            $topic_count =+ 1
        end
        puts "Number of segments #{$num_of_segments}"
        $m_segment_name = Array.new
        $m_ts_data = Array.new
        $m_g_spot = 0
        $m_seg_count = 0
        
        while $m_seg_count < $num_of_segments
            $m_graph_topics_s = Array.new
            $m_graph_topics_s = m_graph_topics.split(",")
            
            
            #clear out array of ts values
            
            
            $ts_unattracted.clear
            $ts_unattracted.push("Unattracted")
            $ts_apathetic.clear
            $ts_apathetic.push("Apathetic")
            $ts_attracted.clear
            $ts_attracted.push("Attracted")
            $ts_impassioned.clear
            $ts_impassioned.push("Impassioned")
            $ts_total.clear
            $ts_total.push("Total")
            #get values for segment
            
            puts " full gdata before #{$m_gdata}"
            
            $m_graph_topics_s.each do |topic|
                puts "topic being looked at #{topic}"
                $m_ts_data = $m_gdata[topic.strip]
                puts "gdata for debuggin #{$m_gdata[topic.strip]}"
                puts " full gdata after #{$m_gdata}"
                puts "ts data #{$m_ts_data}"
                if $m_seg_count == 0
                    $m_g_spot = 1
                    
                    else
                    
                    $m_g_spot = ($m_seg_count * 2) + 1
                    
                end
                
                ts_data_a = $m_ts_data[$m_g_spot.to_i]
                puts "TS Data a for segments #{ts_data_a}"
                
                $ts_unattracted.push(ts_data_a[0])
                $ts_apathetic.push(ts_data_a[1])
                $ts_attracted.push(ts_data_a[2])
                $ts_impassioned.push(ts_data_a[3])
                $ts_total.push(ts_data_a[4])
                
                $name_count = $m_seg_count * 2
                
                $m_segment_name.push($m_ts_data[$name_count.to_i])
                
                $m_segment_name = $m_segment_name.uniq
                
                
                puts "segment name array #{$m_segment_name}"
            end
            $seg_title = "#{gm_topic_code} #{$m_segment_name[$m_seg_count]} - #{gm_topic_title}"
            $seg_title = $seg_title.truncate(30)
            $seg_title = $seg_title.gsub("/", "_")
            
            
            
            #  Per Segments
             puts "Mindset segment Worksheet name #{$seg_title}"
            wb.add_worksheet(:name=> "#{$seg_title}") do |sheet|
                
                sheet.add_row [gm_topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
                
                sheet.add_row [$seg_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
                
                
                #make title value
                if $m_graph_topics_a[0] == ""
                    $m_graph_title = $m_graph_topics_a
                    else
                    $m_graph_title = $m_graph_topics_a.unshift("")
                end
                
                sheet.add_row  $m_graph_title, :style => chart_style
                sheet.add_row  $ts_unattracted , :style => chart_style
                sheet.add_row  $ts_apathetic, :style => chart_style
                sheet.add_row  $ts_attracted, :style => chart_style
                sheet.add_row  $ts_impassioned, :style => chart_style
                sheet.add_row  $ts_total, :style => chart_style_1
                
                
                
                
                num_of_topics = $m_graph_title.size - 1
                
                
                case num_of_topics
                    when 1
                    $t_cells   = "B3:B3"
                    $un_cells  = "B4:B4"
                    $ap_cells  = "B5:B5"
                    $att_cells = "B6:B6"
                    $imp_cells = "B7:B7"
                    when 2
                    $t_cells   = "B3:C3"
                    $un_cells  = "B4:C4"
                    $ap_cells  = "B5:C5"
                    $att_cells = "B6:C6"
                    $imp_cells = "B7:C7"
                    when 3
                    $t_cells   = "B3:D3"
                    $un_cells  = "B4:D4"
                    $ap_cells  = "B5:D5"
                    $att_cells = "B6:D6"
                    $imp_cells = "B7:D7"
                    when 4
                    $t_cells   = "B3:E3"
                    $un_cells  = "B4:E4"
                    $ap_cells  = "B5:E5"
                    $att_cells = "B6:E6"
                    $imp_cells = "B7:E7"
                    when 5
                    $t_cells   = "B3:F3"
                    $un_cells  = "B4:F4"
                    $ap_cells  = "B5:F5"
                    $att_cells = "B6:F6"
                    $imp_cells = "B7:F7"
                    when 6
                    $t_cells   = "B3:G3"
                    $un_cells  = "B4:G4"
                    $ap_cells  = "B5:G5"
                    $att_cells = "B6:G6"
                    $imp_cells = "B7:G7"
                    when 7
                    $t_cells   = "B3:H3"
                    $un_cells  = "B4:H4"
                    $ap_cells  = "B5:H5"
                    $att_cells = "B6:H6"
                    $imp_cells = "B7:H7"
                    when 8
                    $t_cells   = "B3:I3"
                    $un_cells  = "B4:I4"
                    $ap_cells  = "B5:I5"
                    $att_cells = "B6:I6"
                    $imp_cells = "B7:I7"
                    when 9
                    $t_cells   = "B3:J3"
                    $un_cells  = "B4:J4"
                    $ap_cells  = "B5:J5"
                    $att_cells = "B6:J6"
                    $imp_cells = "B7:J7"
                    when 10
                    $t_cells   = "B3:K3"
                    $un_cells  = "B4:K4"
                    $ap_cells  = "B5:K5"
                    $att_cells = "B6:K6"
                    $imp_cells = "B7:K7"
                    when 11
                    $t_cells   = "B3:L3"
                    $un_cells  = "B4:L4"
                    $ap_cells  = "B5:L5"
                    $att_cells = "B6:L6"
                    $imp_cells = "B7:L7"
                    when 12
                    $t_cells   = "B3:M3"
                    $un_cells  = "B4:M4"
                    $ap_cells  = "B5:M5"
                    $att_cells = "B6:M6"
                    $imp_cells = "B7:M7"
                    when 13
                    $t_cells   = "B3:N3"
                    $un_cells  = "B4:N4"
                    $ap_cells  = "B5:N5"
                    $att_cells = "B6:N6"
                    $imp_cells = "B7:N7"
                    when 14
                    $t_cells   = "B3:O3"
                    $un_cells  = "B4:O4"
                    $ap_cells  = "B5:O5"
                    $att_cells = "B6:O6"
                    $imp_cells = "B7:O7"
                    when 15
                    $t_cells   = "B3:P3"
                    $un_cells  = "B4:P4"
                    $ap_cells  = "B5:P5"
                    $att_cells = "B6:P6"
                    $imp_cells = "B7:P7"
                    else
                    puts "Out of bounds, double check number of permitted topics per graph."
                end
                
                
                
                
                sheet.add_chart(Axlsx::Bar3DChart, :start_at => "B10", :end_at => "O32",:title=>"   ", :show_legend => false, :barDir => :bar, :grouping => :percentStacked, ) do |chart|
                    chart.add_series :data => sheet[$un_cells], :colors => ['be0712', 'be0712', 'be0712'], :labels => sheet[$t_cells], :color => "FFFFFF"
                    chart.add_series :data => sheet[$ap_cells], :fg_color => "ffffff" , :colors => ['a5a5a5', 'a5a5a5', 'a5a5a5']
                    chart.add_series :data => sheet[$att_cells], :fg_color => "ffffff" , :colors => ['ffffff', 'ffffff', 'ffffff']
                    chart.add_series :data => sheet[$imp_cells], :fg_color => "ffffff" , :colors => ['365e92', '365e92', '365e92']
                    chart.d_lbls.show_val = true
                    chart.d_lbls.show_percent = true
                    chart.valAxis.gridlines = true
                    chart.catAxis.gridlines = false
                    chart.valAxis.format_code = "Percentage"
                    
                    $m_seg_count += 1
                    
                end
                
            end
            
            
        end
    end


 end #end values.each |do|
    p.serialize($file_name)
    file = $file_name.to_s
    FileUtils.mv $file_name, "./public/uploads/#{$file_name}/graphs_#{$file_name}"




end

#makeGraph("P1", "standard_3x", "Q378", "Role in Caring for Patient" , "When I think about my role in caring for my patients with Pseudobulbar affect (PBA), I feel:", "SEG", "Q151")




end

end#end parser

