require 'creek'
require 'axlsx'
require 'axlsx_rails'
require 'rubygems'





#Parse doc incoming
cworkbook = Creek::Book.new 'sample_mapping_5.xlsx'
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
    
    



#######################################
######## make verbatim start ##########
#######################################
#def makeVerb(topic_code, topic_type, survey_column, topic_title, topic_frame_of_reference, segment, product_mindset , ranking_num)
def makeVerb(values)

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




distance_counter = 0
distance = 3

distance_row = $data[0].to_s.split("   ")
start_spot = distance_row.index(survey_column)
$g_start_spot = distance_row.index(survey_column)
segment_calc = distance_row.index(segment)
product_start_spot = distance_row.index(product_mindset)
$rank_num = ranking_num





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
    #statement_rank = number of statements
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
    while count < statement_rank
        
       r_calc_column = $g_start_spot + count
       
       r_calc_value = data_row[r_calc_column]
       
       
       if r_calc_value.to_i <= statement_count.to_i
           calc_total = data_row[r_calc(r_calc_value, statement_rank , "valence" )].to_i + calc_total
        
       end
       count = count + 1
    end
    
    mindsetc = calc_total / statement_count
    
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
   
   #Ranking
   when "ranking"
   $data.each do |row|
      data_row = $data[row_count].to_s.split("   ") # Parses individual rows from sheet
      if ranking_num.to_i < data_row[s11].to_i
          puts "Not evaluated statement, skipping"
          row_count = row_count + 1
          else
       if row == $data[0]
           
           puts "Skipping Main Row"
           row_count = row_count + 1
           elsif row == $data[1]
           
           puts "Skipping second row"
           row_count = row_count + 1
           else
           
         
           
           sheet.add_row [data_row[r_calc(data_row[s11], 8 , "emotion" )] , data_row[r_calc(data_row[s11], 8 , "intensity" )], data_row[r_calc(data_row[s11], 8 , "why" ) ], "", valence_calc(data_row[r_calc(data_row[s11], 8 , "valence" )]), r_mindset_calc( 8, $rank_num, row_count), data_row[segment_calc] , mindset_calc(data_row[p14],data_row[p24],data_row[p34]) , data_row[0] ], :sz => 10, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
           
           
           row_count = row_count + 1
       end
    
       
   end
end
    else puts "Not a known topic code"

end # end topic type case

#sheet.auto_filter = "A3:I52"
sheet.add_table "A3:I10"
 end # end worksheet creator


end # end while to catch all sheets

p.serialize('test.xlsx')
 end



#######################################
########## make graph start ###########
#######################################

def makeGraph(graph_code, topic_type, survey_column, topic_title, topic_frame_of_reference, segment, product_mindset )

p = Axlsx::Package.new
wb = p.workbook


distance_counter = 0
distance = 3

distance_row = $data[0].to_s.split("   ")
start_spot = distance_row.index(survey_column)
segment_calc = distance_row.index(segment)
product_start_spot = distance_row.index(product_mindset)



$s11 = start_spot       # Emotion
$s12 = start_spot + 4   # Intensity
$s13 = start_spot + 5   # Why
$s14 = start_spot + 6   # Valence
$s21 = start_spot + 1   # Emotion
$s22 = start_spot + 8   # Intensity
$s23 = start_spot + 9   # Why
$s24 = start_spot + 10  # Valence
$s31 = start_spot + 2   # Emotion
$s32 = start_spot + 12  # Intensity
$s33 = start_spot + 13  # Why
$s34 = start_spot + 14  # Valence

$p14 = product_start_spot + 6   # Valence
$p24 = product_start_spot + 10  # Valence
$p34 = product_start_spot + 14  # Valence



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

def positive_calc()
    positive = 0
    neutral = 0
    negative = 0
    
    
    row_count = 0
    
   
    $data.each do |row|
        if row == $data[0]
            
            row_count = row_count + 1
            elsif row == $data[1]
            
            row_count = row_count + 1
            else
            
            $data_g_row = $data[row_count].to_s.split("   ") # Parses individual rows from sheet
        
            if $data_g_row[$s14].to_i == 1
                positive = positive + 1
              elsif $data_g_row[$s14].to_i == 2
                positive = positive + 1
              elsif $data_g_row[$s14].to_i == 3
                neutral = neutral + 1
              elsif $data_g_row[$s14].to_i == 4
                negative = negative + 1
              elsif $data_g_row[$s14].to_i == 5
                negative = negative + 1
             end
            if $data_g_row[$s24].to_i == 1
                positive = positive + 1
                elsif $data_g_row[$s24].to_i == 2
                positive = positive + 1
                elsif $data_g_row[$s24].to_i == 3
                neutral = neutral + 1
                elsif $data_g_row[$s24].to_i == 4
                negative = negative + 1
                elsif $data_g_row[$s24].to_i == 5
                negative = negative + 1
            end
            if $data_g_row[$s34].to_i == 1
                positive = positive + 1
                elsif $data_g_row[$s34].to_i == 2
                positive = positive + 1
                elsif $data_g_row[$s34].to_i == 3
                neutral = neutral + 1
                elsif $data_g_row[$s34].to_i == 4
                negative = negative + 1
                elsif $data_g_row[$s34].to_i == 5
                negative = negative + 1
            end
            
        
            row_count = row_count + 1
        end
        
    end
  
    $total = positive + neutral + negative
    $per_positive = (positive.to_f / $total.to_f).round(2)
    $per_neutral  = (neutral.to_f  / $total.to_f).round(2)
    $per_negative = (negative.to_f / $total.to_f).round(2)
    puts "per positive #{$per_positive}"
    puts "per negative #{$per_negative}"
    puts "per neutral  #{$per_neutral}"
end




row_count = 0


wb.add_worksheet(:name=> "#{graph_code} - #{topic_title}") do |sheet|
    
    sheet.add_row [topic_title], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    sheet.add_row [topic_frame_of_reference], :sz => 12, :b => true, :alignment => { :horizontal => :center, :vertical => :center , :wrap_text => true}
    
    

            positive_calc()
            sheet.add_row ["Condition"]
            sheet.add_row ["Positive" , $per_positive]
            sheet.add_row ["Neutral"  , $per_neutral]
            sheet.add_row ["Negative" , $per_negative]
            sheet.add_row ["Total"    , $total]
            
      
      
      %w(positive neutral negative).each #{ |label| sheet.add_row [label, rand(24)+1] }
      sheet.add_chart(Axlsx::Bar3DChart, :start_at => "D9", :end_at => "H21",  :barDir => :col) do |chart|
          chart.add_series :data => sheet["B4:B6"], :labels => sheet["A4:A6"], :colors => ['FF0000', '00FF00', '0000FF']
      end


    
    
    p.serialize('test_graph.xlsx')
end

end

#makeVerb("R1", "ranking", "Q231_9", "Attributes" , "NUEDEXTA is the first and only FDA-approved treatment for PBA.", "SEG", "Q285", 5)

#makeGraph("P1", "standard_3x", "Q378", "Role in Caring for Patient" , "When I think about my role in caring for my patients with Pseudobulbar affect (PBA), I feel:", "SEG", "Q151")





makeVerb([["R1", "ranking", "Q231_9", "Attributes" , "NUEDEXTA is the first and only FDA-approved treatment for PBA.", "SEG", "Q285", 5],["T1", "standard_3x", "Q24", "Impact of My Alz" , "When I think about the impact of Alzheimer’s/Dementia, I feel:", "SEG", "Q285", 5]])


end


