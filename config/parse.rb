require 'creek'
require 'axlsx'
require 'axlsx_rails'
require 'rubygems'
require 'fileutils'





class Parser
 
def initialize(mapping)
#Parse doc incoming
$file_name = mapping[1].to_s
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
        $data.push(row.values.join "   ")
        # uncomment to print out row values
        #puts row_cells.join "   "
        
    end
  
   
    data_col = Array.new
    col_one = Array.new
    
    


end
#######################################
######## make verbatim start ##########
#######################################
#def makeVerb(topic_code, topic_type, survey_column, topic_title, topic_frame_of_reference, segment, product_mindset , ranking_num, ranking_total)
def makeVerb( values )

run_count = values.size

p = Axlsx::Package.new
wb = p.workbook

values.each do |verbatim_call|
    
topic_code               = verbatim_call[0]
topic_type               = verbatim_call[1]
survey_column            = verbatim_call[2]
topic_title              = verbatim_call[3]
topic_frame_of_reference = verbatim_call[4]
segment                  = verbatim_call[5]
product_mindset          = verbatim_call[6]
ranking_num              = verbatim_call[7]
ranking_total            = verbatim_call[8]




distance_counter = 0
distance = 3

distance_row = $data[0].to_s.split("   ")
start_spot = distance_row.index(survey_column)
$g_start_spot = distance_row.index(survey_column)
segment_calc = distance_row.index(segment)
product_start_spot = distance_row.index(product_mindset)
$rank_num = ranking_num
$rank_total = ranking_total




def valence_calc(value)
   case value
    when "1"
      return "Very Pleasant"
    when "2"
      return "Mildly Pleasant"
    when "3"
      return "Neutral"
    when "4"
      return "Mildly Unpleasant"
    when "5"
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
    end
    #leave that 6 in there, it corresponds to how the mapping files are setup
    find_att   = (col_val.to_i * 6) - 6
    att_offset = statement_rank.to_i + $g_start_spot
    
    
    return find_att + att_offset + type_val

end


def r_mindset_calc( statement_rank, statement_count, row_num)
    data_row = $data[row_num].to_s.split("   ") # Parses individual rows from sheet
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


 wb.add_worksheet(:name=> "#{topic_code} - #{topic_title}") do |sheet|
     
     sheet.add_row [topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}

     sheet.add_row [topic_frame_of_reference], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
     sheet.add_row ["Emotion", "Intensity", "Why", "S" , "Valence", "Mindset", "Segment", "Product Mindset" ,"Response ID"]
     
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
        data_row = $data[row_count].to_s.split("   ") # Parses individual rows from sheet
                              # increments count
        
        
        
        sheet.add_row [ data_row[s11] , data_row[s12] , data_row[s13] , data_row[s14] , valence_calc(data_row[s14]), mindset_calc(data_row[s14],data_row[s24],data_row[s34]), data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
        
        sheet.add_row [ data_row[s21] , data_row[s22] , data_row[s23] , data_row[s24] , valence_calc(data_row[s24]), mindset_calc(data_row[s14],data_row[s24],data_row[s34]), data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}

        sheet.add_row [ data_row[s31] , data_row[s32] , data_row[s33] , data_row[s34] , valence_calc(data_row[s34]), mindset_calc(data_row[s14],data_row[s24],data_row[s34]), data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}

        row_count = row_count + 1
       end
       
    end
   pop_row_count = (row_count * 3) - 2
   
   
   
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
           data_row = $data[row_count].to_s.split("   ") # Parses individual rows from sheet
           # increments count
           
           
           
           sheet.add_row [ data_row[ss11] , data_row[ss12] , data_row[ss13] , data_row[ss14] , valence_calc(data_row[ss14]), data_row[s14], data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
           
           
           
           row_count = row_count + 1
       end
       
   end
   pop_row_count = (row_count * 3) - 2
   
   
   #Ranking
   when "ranking"
   original_start = s11
   
   while s11 <= original_start + $rank_total.to_i

   if s11 == original_start
       
       $data.each do |row|
           data_row = $data[row_count].to_s.split("   ") # Parses individual rows from sheet
           
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
               
               
               
               sheet.add_row [data_row[r_calc(data_row[s11], $rank_total , "emotion" )] , data_row[r_calc(data_row[s11], $rank_total , "intensity" )], data_row[r_calc(data_row[s11], $rank_total , "why" ) ], "", valence_calc(data_row[r_calc(data_row[s11], $rank_total , "valence" )]), r_mindset_calc( $rank_total, $rank_num, row_count), data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
               
               
               row_count = row_count + 1
           end
       end
       
   else
   row_count = 0
   wb.add_worksheet(:name=> "#{topic_code} - #{topic_title}_#{s11}") do |sheet|
       
       sheet.add_row [topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
       
       sheet.add_row [topic_frame_of_reference], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
       sheet.add_row ["Emotion", "Intensity", "Why", "S" , "Valence", "Mindset", "Segment", "Product Mindset" ,"Response ID"]
       
       
       $data.each do |row|
           data_row = $data[row_count].to_s.split("   ") # Parses individual rows from sheet
           
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
               
               
               
               sheet.add_row [data_row[r_calc(data_row[s11], $rank_total , "emotion" )] , data_row[r_calc(data_row[s11], $rank_total , "intensity" )], data_row[r_calc(data_row[s11], $rank_total , "why" ) ], "", valence_calc(data_row[r_calc(data_row[s11], $rank_total , "valence" )]), r_mindset_calc( $rank_total, $rank_num, row_count), data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
               
               
               row_count = row_count + 1
           end
       end
       
       
       
   end

   

  end
   s11 = s11 + 1
 end
   
   
   
   
   else puts "Not a known topic code"




end # end topic type case

#sheet.auto_filter = "A3:I52"
sheet.add_table "A3:I10"
 end # end worksheet creator
#test commit number 2


end # end while to catch all sheets

p.serialize($file_name)
file = $file_name.to_s
FileUtils.mv $file_name, "./public/uploads/#{$file_name}/verbatim_#{$file_name}"
end



#######################################
########## make graph start ###########
#######################################

#def makeGraph(g_topic_code, g_topic_type, graph_topics, g_topic_title, g_ranking_num, g_ranking_total )
def makeGraph(values)
    
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
    $g_data_raw = Array.new
    #################################   Percent Positive Calcs   ###############################
    #Parse verbatim by sheet
    cworksheets.each do |cworksheet|
        puts "Reading verbatims for graph generation, sheet: #{cworksheet.name}"
        num_rows = 0
    
        #get the contents of the sheet and put them in an array
        cworksheet.rows.each do |row|
            row_cells = row.values
            num_rows += 1
            $g_data_raw.push(row.values.join "   ")
        end
        #Percent Positive start values
        row_count = 0
        positive = 0
        neutral = 0
        negative = 0
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
                data_row = $g_data_raw[row_count].to_s.split("   ") # Parses individual rows from sheet
                
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
                #add segments to array so you can manipulate it later
                
                $segment_a.push(data_row[6])
                
                
                row_count = row_count + 1
            end
            
            
        end # End for parsing calculation
        $total = positive + neutral + negative
        $per_positive = (positive.to_f / $total.to_f).round(2)
        $per_neutral  = (neutral.to_f  / $total.to_f).round(2)
        $per_negative = (negative.to_f / $total.to_f).round(2)
        puts "per positive #{$per_positive}"
        puts "per negative #{$per_negative}"
        puts "per neutral  #{$per_neutral}"
        
        
        
    
        #################################   Segment Calcs   ###############################
        #Run segment calcs, this is all dynamic and should work with any number of segments
        row_count       = 0
        $seg_array      = Array.new
        $uniq_seg_names = Array.new
        $uniq_segs      = Array.new
        
        
        #how many segments are they and what are the unique vals
        $num_of_segs     = $segment_a.uniq.length
        $uniq_seg_names  = $segment_a.uniq
        
        
        #turn segs values into a nested array
        $uniq_seg_names.each do |seg|
            array_spot = $uniq_seg_names.index(seg)
            $uniq_segs[array_spot] = Array.new
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
                data_row = $g_data_raw[row_count].to_s.split("   ") # Parses individual rows from sheet
                
             
                $uniq_seg_names.each do |seg| # see what segment is on the row we're looking at
                    array_spot = $uniq_seg_names.index(seg)
                    if seg == data_row[6]
                        
                       $uniq_segs[array_spot].push(data_row[4]) # Put the valence into an array corresponding to the segment title
                    end
                
               end
                row_count = row_count + 1
            end
            
      end # End segment row parsing
      
            $seg_values = Hash.new
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
        $seg_values = $seg_values.reject { |k,v| k.nil? }
        puts "Debugging seg_values hash, #{$seg_values}"
        $ts = [ $per_positive, $per_neutral, $per_negative,  $total, $total]
        title = cworksheet.name[0..2].strip
        $gdata[title] = [$ts]
        
       
        ##TODO - Filter out nil values from
         $seg_values.each do |key, value|
             title = cworksheet.name[0..2].strip
             $gdata[title].push(key, value)
             
        end
         
      
      end # End for each worksheet
    puts $gdata
    #create new xlsx doc
    p = Axlsx::Package.new
    wb = p.workbook
    
    
    
    values.each do |graph_call|
        
        g_topic_code          = graph_call[0]
        g_topic_type          = graph_call[1]
        graph_topics          = graph_call[2]
        g_topic_title         = graph_call[3]
        g_ranking_num         = graph_call[4]
        g_ranking_total       = graph_call[5]
        
        $graph_topics_a = Array.new
        $graph_topics_a = graph_topics.split(",")


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
    ts_data = $gdata[topic.strip]
    ts_data_a = ts_data[0]
    
    $ts_positive.push(ts_data_a[0])
    $ts_neutral.push(ts_data_a[1])
    $ts_negative.push(ts_data_a[2])
    $ts_total_used.push(ts_data_a[3])
    $ts_total.push(ts_data_a[4])
    
end



#  TS
wb.add_worksheet(:name=> "#{g_topic_code} (TS) - #{g_topic_title}") do |sheet|
    
    sheet.add_row [g_topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    sheet.add_row [g_topic_type], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}


            sheet.add_row  $graph_topics_a.unshift("")
            sheet.add_row  $ts_positive #positive
            sheet.add_row  $ts_neutral #neutral
            sheet.add_row  $ts_negative #negative
            sheet.add_row  $ts_total #total
            sheet.add_row  $ts_total_used #message index
            
            num_of_topics = $graph_topics_a.size - 1
            case num_of_topics
                when 1
                   $p_cells   = "B4"
                   $neu_cells = "B5"
                   $neg_cells = "B6"
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
      
     
          sheet.add_chart(Axlsx::Bar3DChart, :start_at => "D9", :end_at => "H21",  :barDir => :col, :grouping => :percentStacked) do |chart|
          chart.add_series :data => sheet[$p_cells], :title => "Positive", :colors => ['00FF00', '00FF00', '00FF00']
          chart.add_series :data => sheet[$neu_cells], :title => "Neutral", :colors => ['0000FF', '0000FF', '0000FF']
          chart.add_series :data => sheet[$neg_cells], :title => "Negative", :colors => ['FF0000', 'FF0000', 'FF0000']
          # chart.add_series :data => sheet["B5:C5"], :labels => sheet["A4:A6"], :colors => ['00FF00', '0000FF', 'FF0000']
      end




end


#  Per Segments
wb.add_worksheet(:name=> "#{g_topic_code} 2 - #{g_topic_title}") do |sheet|
    
    sheet.add_row [g_topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    sheet.add_row ["test2"], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    
    
    
    sheet.add_row ["Condition"]
    sheet.add_row ["Positive" , $per_positive]
    sheet.add_row ["Neutral"  , $per_neutral]
    sheet.add_row ["Negative" , $per_negative]
    sheet.add_row ["Total"    , $total]
    
    
    
   
    #  sheet.add_chart(Axlsx::Bar3DChart, :start_at => "D9", :end_at => "H21",  :barDir => :col) do |chart|
    #     chart.add_series :data => sheet["B4:B6"], :labels => sheet["A4:A6"], :colors => ['FF0000', '00FF00', '0000FF']
    #end
    
    
    
    
end







    #p.serialize('test_graph.xlsx')
  
end
    p.serialize($file_name)
    file = $file_name.to_s
    FileUtils.mv $file_name, "./public/uploads/#{$file_name}/graphs_#{$file_name}"

end

end


#makeGraph("P1", "standard_3x", "Q378", "Role in Caring for Patient" , "When I think about my role in caring for my patients with Pseudobulbar affect (PBA), I feel:", "SEG", "Q151")




end#end parser
#end
