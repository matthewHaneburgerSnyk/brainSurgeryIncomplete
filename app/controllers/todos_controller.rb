require 'rubygems'
require 'json'
require 'carrierwave'
require 'carrierwave/orm/activerecord'
require '././config/parse.rb'
require '././config/export.rb'

class TodosController < ApplicationController
  
  
  before_action :set_todo, only: [:show, :edit, :update, :destroy, :download_verbatim]

  # GET /todos
  # GET /todos.json
  def index
    @todos = Todo.all
  end

  # GET /todos/1
  # GET /todos/1.json
  def show
  end

  # GET /todos/new
  def new
    @todo = Todo.new
  end

  # GET /todos/1/edit
  def edit
      
  end

  # POST /todos
  # POST /todos.json
  def save
      @export = Export.new(todo_params[:project_name])
     
     @todo = Todo.new(todo_params)
     filename = todo_params[:mapping_file].original_filename
     filename = filename.gsub(" ", "_")
     @todo.mapping_file_name = filename
     
     respond_to do |format|
         if @todo.save
             format.html { redirect_to @todo, notice: 'Project was successfully created.' }
             format.json { render :show, status: :created, location: @todo }
             uploader = MappingUploader.new
             uploader.store!(todo_params[:mapping_file])
             
             todosV           = v_todo_params
             todosVH          = eval(todosV.to_s)
             todosVH          = todosVH.collect { |k, v| v }
             todosVH          = todosVH.each_slice(8).to_a
             #You're not always going to have 7 of these ^ figure out how to account for that
             todosMinds        = Array.new
             todosMindsTypes   = Array.new
             todosMindsTitles  = Array.new
             
             todosVH.each do |values|
                 puts "todos VH values #{values}"
                 count = 0
                 if values[7].to_i == 1
                     puts "minds value #{values[2]}"
                     puts "types value #{values[1]}"
                     puts "titles value #{values[3]}"
                     
                     
                     todosMinds.push(values[2])
                     todosMindsTypes.push(values[1])
                     todosMindsTitles.push(values[3])
                     
                 end
                 count += 1
             end

            segs             = segment_params
            todosSegs        = eval(segs.to_s)
            todosSegs        = todosSegs.collect { |k, v| v}

            todosG  = g_todo_params
            todosGH = eval(todosG.to_s)
            todosGH = todosGH.collect { |k, v| v }
            todosGH = todosGH.each_slice(6).to_a


             
             
             
             
             
             @export.makeExport(todosVH, todosGH,todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
             else
             format.html { render :new }
             format.json { render json: @todo.errors, status: :unprocessable_entity }
         end
     end
  end
  
  
  def create
   @filename = ""
     #create new todo
   @todo = Todo.new(todo_params)
   if todo_params[:mapping_file].nil?
       puts "Skipping submit, no file present"
   else
   
   filename = todo_params[:mapping_file].original_filename
   filename = filename.gsub(" ", "_")
   filename = filename.gsub("(", "")
   filename = filename.gsub(")", "")
   filename = filename.gsub("/", "")
   @todo.mapping_file_name = filename#todo_params[:mapping_file].original_filename

   
   end
   
   #Generate json/html and insert all the form values into the DB
    respond_to do |format|
      if @todo.save
          
        if todo_params[:mapping_file].nil?
            
            puts "Skipping file save, no file present"
           else
            @filename = todo_params[:mapping_file].original_filename
            @filename = @filename.gsub(" ", "_")
            @filename = @filename.gsub("(", "")
            @filename = @filename.gsub(")", "")
            @filename = @filename.gsub("/", "")
            @todo.mapping_file_name = @filename
            uploader = MappingUploader.new
            uploader.store!(todo_params[:mapping_file])
            puts"This is the filename >>>> #{@filename}"
            
        end
        
        
        
        format.html { redirect_to @todo, notice: 'Project was successfully created.' }
        format.json { render :show, status: :created, location: @todo }

        
      else
        format.html { render :new }
        format.json { render json: @todo.errors, status: :unprocessable_entity }
      end
    end
    #Get original filename and put it where it goes
    
    if params[:commit] == "Save"
        puts "Skipping calcs since this is a save"
        
        
        @export          = Export.new(todo_params[:project_name])
        todosV           = v_todo_params
        todosVH          = eval(todosV.to_s)
        todosVH          = todosVH.collect { |k, v| v }
        todosVH          = todosVH.each_slice(8).to_a
        
        todosMinds        = Array.new
        todosMindsTypes   = Array.new
        todosMindsTitles  = Array.new
        
        todosVH.each do |values|
            puts "todos VH values #{values}"
            count = 0
            if values[7].to_i == 1
                puts "minds value #{values[2]}"
                puts "types value #{values[1]}"
                puts "titles value #{values[3]}"
                
                
                todosMinds.push(values[2])
                todosMindsTypes.push(values[1])
                todosMindsTitles.push(values[3])
                
            end
            count += 1
        end
        
        segs             = segment_params
        todosSegs        = eval(segs.to_s)
        todosSegs        = todosSegs.collect { |k, v| v}
        
        todosG  = g_todo_params
        todosGH = eval(todosG.to_s)
        todosGH = todosGH.collect { |k, v| v }
        todosGH = todosGH.each_slice(6).to_a
        
        @export.makeExport(todosVH, todosGH,todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
        

    elsif params[:commit] == "Submit"
   
    puts "mapping filename #{@todo.mapping_file_name}"
    foldername = @filename.split(".")[0]
    #put mapping file where it goes
    map = File.open("././public/uploads/#{@filename}/#{@filename}")
    
    verbatim_path = "././public/uploads/#{@filename}/Verbatim_#{@filename}"
    graph_path    = "././public/uploads/#{@filename}/Graph_#{@filename}"
   
    p_data = [map, @filename]
    
  
    
    
    #new Parser and generate verbatim file
    @parser          = Parser.new(p_data)
    @export          = Export.new(todo_params[:project_name])
    todosV           = v_todo_params
    todosVH          = eval(todosV.to_s)
    todosVH          = todosVH.collect { |k, v| v }
    todosVH          = todosVH.each_slice(8).to_a
    
    
    todosMinds        = Array.new
    todosMindsTypes   = Array.new
    todosMindsTitles  = Array.new
    
    #get mindsets
    
    todosVH.each do |values|
        puts "todos VH values #{values}"
        count = 0
        if values[7].to_i == 1
            puts "minds value #{values[2]}"
            puts "types value #{values[1]}"
            puts "titles value #{values[3]}"
            
            
            todosMinds.push(values[2])
            todosMindsTypes.push(values[1])
            todosMindsTitles.push(values[3])
            
        end
        count += 1
    end
    
    if todosMinds.size > 5
        raise Exception.new('Only 5 mindsets can be submitted.  Please uncheck one row as a mindset.')
        
    end
    
    segs             = segment_params
    todosSegs        = eval(segs.to_s)
    todosSegs        = todosSegs.collect { |k, v| v}
    
    #minds            = mindset_params
    #todosMinds       = eval(minds.to_s)
    #todosMinds       = todosMinds.collect { |k, v| v}
    
    #mind_types       = mindset_params_types
    #todosMindsTypes  = eval(mind_types.to_s)
    #todosMindsTypes  = todosMindsTypes.collect { |k, v| v}
    
    #mind_titles      = mindset_param_titles
    #todosMindsTitles = eval(mind_titles.to_s)
    #todosMindsTitles = todosMindsTitles.collect { |k, v| v}
    
    
    
    
    
    
    @parser.makeVerb(todosVH, todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
    
    #Generate graphs with existing parser
    
    
    todosG  = g_todo_params
    todosGH = eval(todosG.to_s)
    todosGH = todosGH.collect { |k, v| v }
    todosGH = todosGH.each_slice(6).to_a
    puts "todosGH #{todosGH}"
    @parser.makeGraph(todosGH, todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)

    @export.makeExport(todosVH, todosGH,todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
    #puts todosVH.class
    #puts todosVH
    end
    
  end

  # PATCH/PUT /todos/1
  # PATCH/PUT /todos/1.json
  def update
      @filename = ""
      puts "this is todo params #{todo_params}"
      puts "this is params #{params}"
      
      if todo_params[:mapping_file].nil?
          puts "Skipping submit, no file present"
          else
          
          @filename = todo_params[:mapping_file].original_filename
          @filename = @filename.gsub(" ", "_")
          @filename = @filename.gsub("(", "")
          @filename = @filename.gsub(")", "")
          @filename = @filename.gsub("/", "")
          @todo.mapping_file_name = @filename#todo_params[:mapping_file].original_filename
          puts "mapping file name #{@todo.mapping_file_name}"
          puts "original file name #{todo_params[:mapping_file].original_filename}"
          
      end
      
      
      
      respond_to do |format|
      if @todo.update(todo_params)
          
        puts "params mapping file - #{todo_params[:mapping_file]}"
        if todo_params[:mapping_file].nil?
            
         puts "Skipping save, no file present"
         
        else
        #@filename = todo_params[:mapping_file].original_filename
        #@filename = @filename.gsub(" ", "_")
        #@filename = @filename.gsub("(", "")
        #@filename = @filename.gsub(")", "")
        #@filename = @filename.gsub("/", "")
        #@todo.mapping_file_name = @filename
        uploader = MappingUploader.new
        uploader.store!(todo_params[:mapping_file])
        end
        
        
        
        format.html { redirect_to @todo, notice: 'Project was successfully updated.' }
        format.json { render :show, status: :ok, location: @todo }
        
      else
        format.html { render :edit }
        format.json { render json: @todo.errors, status: :unprocessable_entity }
      end
    end
   
    
    if params[:commit] == "Save"
        puts "Skipping calcs since this is a save"
        
        @export          = Export.new(todo_params[:project_name])
        todosV           = v_todo_params
        todosVH          = eval(todosV.to_s)
        todosVH          = todosVH.collect { |k, v| v }
        todosVH          = todosVH.each_slice(8).to_a
        
        todosMinds        = Array.new
        todosMindsTypes   = Array.new
        todosMindsTitles  = Array.new
        
        todosVH.each do |values|
            puts "todos VH values #{values}"
            count = 0
            if values[7].to_i == 1
                puts "minds value #{values[2]}"
                puts "types value #{values[1]}"
                puts "titles value #{values[3]}"
                
                
                todosMinds.push(values[2])
                todosMindsTypes.push(values[1])
                todosMindsTitles.push(values[3])
                
            end
            count += 1
        end
        
        segs             = segment_params
        todosSegs        = eval(segs.to_s)
        todosSegs        = todosSegs.collect { |k, v| v}
        
        todosG  = g_todo_params
        todosGH = eval(todosG.to_s)
        todosGH = todosGH.collect { |k, v| v }
        todosGH = todosGH.each_slice(6).to_a
        
        
        
        @export.makeExport(todosVH, todosGH,todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
        
        
        



    elsif params[:commit] == "Submit"
    #Get original filename and put it where it goes
    #puts " filename 1 #{@filename}"
    #@filename = todo_params[:mapping_file_name]
    #puts " todo params mapping file name #{params[:mapping_file_name]}  "
    #puts " filename 2 #{@filename}"
    #put mapping file where it goes
    
    map = File.open("././public/uploads/#{@filename}/#{@filename}")
   
    verbatim_path = "././public/uploads/#{@filename}/Verbatim_#{@filename}"
    graph_path    = "././public/uploads/#{@filename}/Graph_#{@filename}"
    
    p_data = [map, @filename]
    
    
    #new Parser and generate verbatim file
    @parser          = Parser.new(p_data)
    @export          = Export.new(todo_params[:project_name])
    todosV           = v_todo_params
    todosVH          = eval(todosV.to_s)
    todosVH          = todosVH.collect { |k, v| v }
    todosVH          = todosVH.each_slice(8).to_a
    
    
    todosMinds        = Array.new
    todosMindsTypes   = Array.new
    todosMindsTitles  = Array.new
    
    #get mindsets
    
    todosVH.each do |values|
        puts "todos VH values #{values}"
        count = 0
        if values[7].to_i == 1
            puts "minds value #{values[2]}"
            puts "types value #{values[1]}"
            puts "titles value #{values[3]}"
            
            
            todosMinds.push(values[2])
            todosMindsTypes.push(values[1])
            todosMindsTitles.push(values[3])
            
        end
        count += 1
    end
    
    if todosMinds.size > 5
       raise Exception.new('Only 5 mindsets can be submitted.  Please uncheck one row as a mindset.')
        
    end
    puts "todos Minds #{todosMinds}"
    
    segs             = segment_params
    todosSegs        = eval(segs.to_s)
    todosSegs        = todosSegs.collect { |k, v| v}
    
    #minds            = mindset_params
    #todosMinds       = eval(minds.to_s)
    #todosMinds       = todosMinds.collect { |k, v| v}
    
    #mind_types       = mindset_params_types
    #todosMindsTypes  = eval(mind_types.to_s)
    #todosMindsTypes  = todosMindsTypes.collect { |k, v| v}
    
    #mind_titles      = mindset_param_titles
    #todosMindsTitles = eval(mind_titles.to_s)
    #todosMindsTitles = todosMindsTitles.collect { |k, v| v}
    
    
    @parser.makeVerb(todosVH, todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
    
    #Generate graphs with existing parser
    
    
    todosG  = g_todo_params
    todosGH = eval(todosG.to_s)
    todosGH = todosGH.collect { |k, v| v }
    todosGH = todosGH.each_slice(6).to_a
    puts "todosGH #{todosGH}"
    @parser.makeGraph(todosGH, todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
    @export.makeExport(todosVH, todosGH,todosSegs, todosMinds, todosMindsTypes, todosMindsTitles)
    end
    
  end

  # DELETE /todos/1
  # DELETE /todos/1.json
  def destroy
    @todo.destroy
    respond_to do |format|
      format.html { redirect_to todos_url, notice: 'Project was successfully destroyed.' }
      format.json { head :no_content }
    end
  end


def delete
    # @todo.delete
    @todo = Todo.find(:all, :conditions => [ "mail = ?", params[:mail]])

end


def download_verbatim

    send_file(
              "#{Rails.root}/public/uploads/#{$filename}/verbatim_#{$filename}",
              filename: "#{$filename}",
              type: "xlsx"
              )
end



  private
  
    # Use callbacks to share common setup or constraints between actions.
    def set_todo
  
            
      @todo = Todo.find(params[:id])
   
    end
    
    def item_params
        params.require(:item).permit(:name, :description ,:mapping_file) # Add :picture as a permitted paramter
    end


    # Never trust parameters from the scary internet, only allow the white list through.
    def mindset_params
        params.require(:todo).permit(:mindset_1, :mindset_2, :mindset_3, :mindset_4, :mindset_5)
     end
    
    def mindset_params_types
        params.require(:todo).permit(:mindset_1_type, :mindset_2_type, :mindset_3_type, :mindset_4_type, :mindset_5_type)
    end
    
    def mindset_param_titles
        params.require(:todo).permit(:mindset_1_title, :mindset_2_title, :mindset_3_title, :mindset_4_title, :mindset_5_title )
    end
    
    
    def segment_params
        params.require(:todo).permit(:segment_1, :segment_2, :segment_3 )
    end
    
    
    
    
    def todo_params
        params.require(:todo).permit(:project_name, :notes, :code, :project_description, :verbatim_file, :graph_file, :mindset_1_title, :mindset_2_title, :mindset_3_title, :mindset_4_title, :mindset_5_title, :mindset_1_type, :mindset_2_type, :mindset_3_type, :mindset_4_type, :mindset_5_type , :segment_1, :segment_2, :segment_3, :mindset_1, :mindset_2, :mindset_3, :mindset_4, :mindset_5, :g_topic_code , :g_topic_type , :g_survey_column , :g_topic_title , :g_topic_frame , :g_segment , :g_product_mindset , :g_ranking_num , :g_topic_code0 , :g_topic_type0 , :g_survey_column0 , :g_topic_title0 , :g_topic_frame0 , :g_segment0 , :g_product_mindset0 , :g_ranking_num0 , :g_topic_code1 , :g_topic_type1 , :g_survey_column1 , :g_topic_title1 , :g_topic_frame1 , :g_segment1 , :g_product_mindset1 , :g_ranking_num1 , :g_topic_code2 , :g_topic_type2 , :g_survey_column2 , :g_topic_title2 , :g_topic_frame2 , :g_segment2 , :g_product_mindset2 , :g_ranking_num2 , :g_topic_code3 , :g_topic_type3 , :g_survey_column3 , :g_topic_title3 , :g_topic_frame3 , :g_segment3 , :g_product_mindset3 , :g_ranking_num3 , :g_topic_code4 , :g_topic_type4 , :g_survey_column4 , :g_topic_title4 , :g_topic_frame4 , :g_segment4 , :g_product_mindset4 , :g_ranking_num4 , :g_topic_code5 , :g_topic_type5 , :g_survey_column5 , :g_topic_title5 , :g_topic_frame5 , :g_segment5 , :g_product_mindset5 , :g_ranking_num5 , :g_topic_code6 , :g_topic_type6 , :g_survey_column6 , :g_topic_title6 , :g_topic_frame6 , :g_segment6 , :g_product_mindset6 , :g_ranking_num6 , :g_topic_code7 , :g_topic_type7 , :g_survey_column7 , :g_topic_title7 , :g_topic_frame7 , :g_segment7 , :g_product_mindset7 , :g_ranking_num7 , :g_topic_code8 , :g_topic_type8 , :g_survey_column8 , :g_topic_title8 , :g_topic_frame8 , :g_segment8 , :g_product_mindset8 , :g_ranking_num8 , :g_topic_code9 , :g_topic_type9 , :g_survey_column9 , :g_topic_title9 , :g_topic_frame9 , :g_segment9 , :g_product_mindset9 , :g_ranking_num9 , :g_topic_code10 , :g_topic_type10 , :g_survey_column10 , :g_topic_title10 , :g_topic_frame10 , :g_segment10 , :g_product_mindset10 , :g_ranking_num10 , :g_topic_code11 , :g_topic_type11 , :g_survey_column11 , :g_topic_title11 , :g_topic_frame11 , :g_segment11 , :g_product_mindset11 , :g_ranking_num11 , :g_topic_code12 , :g_topic_type12 , :g_survey_column12 , :g_topic_title12 , :g_topic_frame12 , :g_segment12 , :g_product_mindset12 , :g_ranking_num12 , :g_topic_code13 , :g_topic_type13 , :g_survey_column13 , :g_topic_title13 , :g_topic_frame13 , :g_segment13 , :g_product_mindset13 , :g_ranking_num13 , :g_topic_code14 , :g_topic_type14 , :g_survey_column14 , :g_topic_title14 , :g_topic_frame14 , :g_segment14 , :g_product_mindset14 , :g_ranking_num14 , :g_topic_code15 , :g_topic_type15 , :g_survey_column15 , :g_topic_title15 , :g_topic_frame15 , :g_segment15 , :g_product_mindset15 , :g_ranking_num15 , :g_topic_code16 , :g_topic_type16 , :g_survey_column16 , :g_topic_title16 , :g_topic_frame16 , :g_segment16 , :g_product_mindset16 , :g_ranking_num16 , :g_topic_code17 , :g_topic_type17 , :g_survey_column17 , :g_topic_title17 , :g_topic_frame17 , :g_segment17 , :g_product_mindset17 , :g_ranking_num17 , :g_topic_code18 , :g_topic_type18 , :g_survey_column18 , :g_topic_title18 , :g_topic_frame18 , :g_segment18 , :g_product_mindset18 , :g_ranking_num18 , :g_topic_code19 , :g_topic_type19 , :g_survey_column19 , :g_topic_title19 , :g_topic_frame19 , :g_segment19 , :g_product_mindset19 , :g_ranking_num19 , :g_topic_code20 , :g_topic_type20 , :g_survey_column20 , :g_topic_title20 , :g_topic_frame20 , :g_segment20 , :g_product_mindset20 , :g_ranking_num20 , :g_topic_code21 , :g_topic_type21 , :g_survey_column21 , :g_topic_title21 , :g_topic_frame21 , :g_segment21 , :g_product_mindset21 , :g_ranking_num21 , :g_topic_code22 , :g_topic_type22 , :g_survey_column22 , :g_topic_title22 , :g_topic_frame22 , :g_segment22 , :g_product_mindset22 , :g_ranking_num22 , :g_topic_code23 , :g_topic_type23 , :g_survey_column23 , :g_topic_title23 , :g_topic_frame23 , :g_segment23 , :g_product_mindset23 , :g_ranking_num23 , :g_topic_code24 , :g_topic_type24 , :g_survey_column24 , :g_topic_title24 , :g_topic_frame24 , :g_segment24 , :g_product_mindset24 , :g_ranking_num24 , :g_topic_code25 , :g_topic_type25 , :g_survey_column25 , :g_topic_title25 , :g_topic_frame25 , :g_segment25 , :g_product_mindset25 , :g_ranking_num25 , :g_topic_code26 , :g_topic_type26 , :g_survey_column26 , :g_topic_title26 , :g_topic_frame26 , :g_segment26 , :g_product_mindset26 , :g_ranking_num26 , :g_topic_code27 , :g_topic_type27 , :g_survey_column27 , :g_topic_title27 , :g_topic_frame27 , :g_segment27 , :g_product_mindset27 , :g_ranking_num27 , :g_topic_code28 , :g_topic_type28 , :g_survey_column28 , :g_topic_title28 , :g_topic_frame28 , :g_segment28 , :g_product_mindset28 , :g_ranking_num28 , :g_topic_code29 , :g_topic_type29 , :g_survey_column29 , :g_topic_title29 , :g_topic_frame29 , :g_segment29 , :g_product_mindset29 , :g_ranking_num29 , :g_topic_code30 , :g_topic_type30 , :g_survey_column30 , :g_topic_title30 , :g_topic_frame30 , :g_segment30 , :g_product_mindset30 , :g_ranking_num30 , :g_topic_code31 , :g_topic_type31 , :g_survey_column31 , :g_topic_title31 , :g_topic_frame31 , :g_segment31 , :g_product_mindset31 , :g_ranking_num31 , :g_topic_code32 , :g_topic_type32 , :g_survey_column32 , :g_topic_title32 , :g_topic_frame32 , :g_segment32 , :g_product_mindset32 , :g_ranking_num32 , :g_topic_code33 , :g_topic_type33 , :g_survey_column33 , :g_topic_title33 , :g_topic_frame33 , :g_segment33 , :g_product_mindset33 , :g_ranking_num33 , :g_topic_code34 , :g_topic_type34 , :g_survey_column34 , :g_topic_title34 , :g_topic_frame34 , :g_segment34 , :g_product_mindset34 , :g_ranking_num34 , :g_topic_code35 , :g_topic_type35 , :g_survey_column35 , :g_topic_title35 , :g_topic_frame35 , :g_segment35 , :g_product_mindset35 , :g_ranking_num35 , :g_topic_code36 , :g_topic_type36 , :g_survey_column36 , :g_topic_title36 , :g_topic_frame36 , :g_segment36 , :g_product_mindset36 , :g_ranking_num36 , :g_topic_code37 , :g_topic_type37 , :g_survey_column37 , :g_topic_title37 , :g_topic_frame37 , :g_segment37 , :g_product_mindset37 , :g_ranking_num37 , :g_topic_code38 , :g_topic_type38 , :g_survey_column38 , :g_topic_title38 , :g_topic_frame38 , :g_segment38 , :g_product_mindset38 , :g_ranking_num38 , :g_topic_code39 , :g_topic_type39 , :g_survey_column39 , :g_topic_title39 , :g_topic_frame39 , :g_segment39 , :g_product_mindset39 , :g_ranking_num39 , :g_topic_code40 , :g_topic_type40 , :g_survey_column40 , :g_topic_title40 , :g_topic_frame40 , :g_segment40 , :g_product_mindset40 , :g_ranking_num40 , :g_topic_code41 , :g_topic_type41 , :g_survey_column41 , :g_topic_title41 , :g_topic_frame41 , :g_segment41 , :g_product_mindset41 , :g_ranking_num41 , :g_topic_code42 , :g_topic_type42 , :g_survey_column42 , :g_topic_title42 , :g_topic_frame42 , :g_segment42 , :g_product_mindset42 , :g_ranking_num42 , :g_topic_code43 , :g_topic_type43 , :g_survey_column43 , :g_topic_title43 , :g_topic_frame43 , :g_segment43 , :g_product_mindset43 , :g_ranking_num43 , :g_topic_code44 , :g_topic_type44 , :g_survey_column44 , :g_topic_title44 , :g_topic_frame44 , :g_segment44 , :g_product_mindset44 , :g_ranking_num44 , :g_topic_code45 , :g_topic_type45 , :g_survey_column45 , :g_topic_title45 , :g_topic_frame45 , :g_segment45 , :g_product_mindset45 , :g_ranking_num45 , :g_topic_code46 , :g_topic_type46 , :g_survey_column46 , :g_topic_title46 , :g_topic_frame46 , :g_segment46 , :g_product_mindset46 , :g_ranking_num46 , :g_topic_code47 , :g_topic_type47 , :g_survey_column47 , :g_topic_title47 , :g_topic_frame47 , :g_segment47 , :g_product_mindset47 , :g_ranking_num47 , :g_topic_code48 , :g_topic_type48 , :g_survey_column48 , :g_topic_title48 , :g_topic_frame48 , :g_segment48 , :g_product_mindset48 , :g_ranking_num48 , :g_topic_code49 , :g_topic_type49 , :g_survey_column49 , :g_topic_title49 , :g_topic_frame49 , :g_segment49 , :g_product_mindset49 , :g_ranking_num49 , :v_topic_code , :v_topic_type , :v_survey_column , :v_topic_title , :v_topic_frame , :v_segment , :v_product_mindset , :v_ranking_num , :v_topic_code0 , :v_topic_type0 , :v_survey_column0 , :v_topic_title0 , :v_topic_frame0 , :v_segment0 , :v_product_mindset0 , :v_ranking_num0 , :v_topic_code1 , :v_topic_type1 , :v_survey_column1 , :v_topic_title1 , :v_topic_frame1 , :v_segment1 , :v_product_mindset1 , :v_ranking_num1 , :v_topic_code2 , :v_topic_type2 , :v_survey_column2 , :v_topic_title2 , :v_topic_frame2 , :v_segment2 , :v_product_mindset2 , :v_ranking_num2 , :v_topic_code3 , :v_topic_type3 , :v_survey_column3 , :v_topic_title3 , :v_topic_frame3 , :v_segment3 , :v_product_mindset3 , :v_ranking_num3 , :v_topic_code4 , :v_topic_type4 , :v_survey_column4 , :v_topic_title4 , :v_topic_frame4 , :v_segment4 , :v_product_mindset4 , :v_ranking_num4 , :v_topic_code5 , :v_topic_type5 , :v_survey_column5 , :v_topic_title5 , :v_topic_frame5 , :v_segment5 , :v_product_mindset5 , :v_ranking_num5 , :v_topic_code6 , :v_topic_type6 , :v_survey_column6 , :v_topic_title6 , :v_topic_frame6 , :v_segment6 , :v_product_mindset6 , :v_ranking_num6 , :v_topic_code7 , :v_topic_type7 , :v_survey_column7 , :v_topic_title7 , :v_topic_frame7 , :v_segment7 , :v_product_mindset7 , :v_ranking_num7 , :v_topic_code8 , :v_topic_type8 , :v_survey_column8 , :v_topic_title8 , :v_topic_frame8 , :v_segment8 , :v_product_mindset8 , :v_ranking_num8 , :v_topic_code9 , :v_topic_type9 , :v_survey_column9 , :v_topic_title9 , :v_topic_frame9 , :v_segment9 , :v_product_mindset9 , :v_ranking_num9 , :v_topic_code10 , :v_topic_type10 , :v_survey_column10 , :v_topic_title10 , :v_topic_frame10 , :v_segment10 , :v_product_mindset10 , :v_ranking_num10 , :v_topic_code11 , :v_topic_type11 , :v_survey_column11 , :v_topic_title11 , :v_topic_frame11 , :v_segment11 , :v_product_mindset11 , :v_ranking_num11 , :v_topic_code12 , :v_topic_type12 , :v_survey_column12 , :v_topic_title12 , :v_topic_frame12 , :v_segment12 , :v_product_mindset12 , :v_ranking_num12 , :v_topic_code13 , :v_topic_type13 , :v_survey_column13 , :v_topic_title13 , :v_topic_frame13 , :v_segment13 , :v_product_mindset13 , :v_ranking_num13 , :v_topic_code14 , :v_topic_type14 , :v_survey_column14 , :v_topic_title14 , :v_topic_frame14 , :v_segment14 , :v_product_mindset14 , :v_ranking_num14 , :v_topic_code15 , :v_topic_type15 , :v_survey_column15 , :v_topic_title15 , :v_topic_frame15 , :v_segment15 , :v_product_mindset15 , :v_ranking_num15 , :v_topic_code16 , :v_topic_type16 , :v_survey_column16 , :v_topic_title16 , :v_topic_frame16 , :v_segment16 , :v_product_mindset16 , :v_ranking_num16 , :v_topic_code17 , :v_topic_type17 , :v_survey_column17 , :v_topic_title17 , :v_topic_frame17 , :v_segment17 , :v_product_mindset17 , :v_ranking_num17 , :v_topic_code18 , :v_topic_type18 , :v_survey_column18 , :v_topic_title18 , :v_topic_frame18 , :v_segment18 , :v_product_mindset18 , :v_ranking_num18 , :v_topic_code19 , :v_topic_type19 , :v_survey_column19 , :v_topic_title19 , :v_topic_frame19 , :v_segment19 , :v_product_mindset19 , :v_ranking_num19 , :v_topic_code20 , :v_topic_type20 , :v_survey_column20 , :v_topic_title20 , :v_topic_frame20 , :v_segment20 , :v_product_mindset20 , :v_ranking_num20 , :v_topic_code21 , :v_topic_type21 , :v_survey_column21 , :v_topic_title21 , :v_topic_frame21 , :v_segment21 , :v_product_mindset21 , :v_ranking_num21 , :v_topic_code22 , :v_topic_type22 , :v_survey_column22 , :v_topic_title22 , :v_topic_frame22 , :v_segment22 , :v_product_mindset22 , :v_ranking_num22 , :v_topic_code23 , :v_topic_type23 , :v_survey_column23 , :v_topic_title23 , :v_topic_frame23 , :v_segment23 , :v_product_mindset23 , :v_ranking_num23 , :v_topic_code24 , :v_topic_type24 , :v_survey_column24 , :v_topic_title24 , :v_topic_frame24 , :v_segment24 , :v_product_mindset24 , :v_ranking_num24 , :v_topic_code25 , :v_topic_type25 , :v_survey_column25 , :v_topic_title25 , :v_topic_frame25 , :v_segment25 , :v_product_mindset25 , :v_ranking_num25 , :v_topic_code26 , :v_topic_type26 , :v_survey_column26 , :v_topic_title26 , :v_topic_frame26 , :v_segment26 , :v_product_mindset26 , :v_ranking_num26 , :v_topic_code27 , :v_topic_type27 , :v_survey_column27 , :v_topic_title27 , :v_topic_frame27 , :v_segment27 , :v_product_mindset27 , :v_ranking_num27 , :v_topic_code28 , :v_topic_type28 , :v_survey_column28 , :v_topic_title28 , :v_topic_frame28 , :v_segment28 , :v_product_mindset28 , :v_ranking_num28 , :v_topic_code29 , :v_topic_type29 , :v_survey_column29 , :v_topic_title29 , :v_topic_frame29 , :v_segment29 , :v_product_mindset29 , :v_ranking_num29 , :v_topic_code30 , :v_topic_type30 , :v_survey_column30 , :v_topic_title30 , :v_topic_frame30 , :v_segment30 , :v_product_mindset30 , :v_ranking_num30 , :v_topic_code31 , :v_topic_type31 , :v_survey_column31 , :v_topic_title31 , :v_topic_frame31 , :v_segment31 , :v_product_mindset31 , :v_ranking_num31 , :v_topic_code32 , :v_topic_type32 , :v_survey_column32 , :v_topic_title32 , :v_topic_frame32 , :v_segment32 , :v_product_mindset32 , :v_ranking_num32 , :v_topic_code33 , :v_topic_type33 , :v_survey_column33 , :v_topic_title33 , :v_topic_frame33 , :v_segment33 , :v_product_mindset33 , :v_ranking_num33 , :v_topic_code34 , :v_topic_type34 , :v_survey_column34 , :v_topic_title34 , :v_topic_frame34 , :v_segment34 , :v_product_mindset34 , :v_ranking_num34 , :v_topic_code35 , :v_topic_type35 , :v_survey_column35 , :v_topic_title35 , :v_topic_frame35 , :v_segment35 , :v_product_mindset35 , :v_ranking_num35 , :v_topic_code36 , :v_topic_type36 , :v_survey_column36 , :v_topic_title36 , :v_topic_frame36 , :v_segment36 , :v_product_mindset36 , :v_ranking_num36 , :v_topic_code37 , :v_topic_type37 , :v_survey_column37 , :v_topic_title37 , :v_topic_frame37 , :v_segment37 , :v_product_mindset37 , :v_ranking_num37 , :v_topic_code38 , :v_topic_type38 , :v_survey_column38 , :v_topic_title38 , :v_topic_frame38 , :v_segment38 , :v_product_mindset38 , :v_ranking_num38 , :v_topic_code39 , :v_topic_type39 , :v_survey_column39 , :v_topic_title39 , :v_topic_frame39 , :v_segment39 , :v_product_mindset39 , :v_ranking_num39 , :v_topic_code40 , :v_topic_type40 , :v_survey_column40 , :v_topic_title40 , :v_topic_frame40 , :v_segment40 , :v_product_mindset40 , :v_ranking_num40 , :v_topic_code41 , :v_topic_type41 , :v_survey_column41 , :v_topic_title41 , :v_topic_frame41 , :v_segment41 , :v_product_mindset41 , :v_ranking_num41 , :v_topic_code42 , :v_topic_type42 , :v_survey_column42 , :v_topic_title42 , :v_topic_frame42 , :v_segment42 , :v_product_mindset42 , :v_ranking_num42 , :v_topic_code43 , :v_topic_type43 , :v_survey_column43 , :v_topic_title43 , :v_topic_frame43 , :v_segment43 , :v_product_mindset43 , :v_ranking_num43 , :v_topic_code44 , :v_topic_type44 , :v_survey_column44 , :v_topic_title44 , :v_topic_frame44 , :v_segment44 , :v_product_mindset44 , :v_ranking_num44 , :v_topic_code45 , :v_topic_type45 , :v_survey_column45 , :v_topic_title45 , :v_topic_frame45 , :v_segment45 , :v_product_mindset45 , :v_ranking_num45 , :v_topic_code46 , :v_topic_type46 , :v_survey_column46 , :v_topic_title46 , :v_topic_frame46 , :v_segment46 , :v_product_mindset46 , :v_ranking_num46 , :v_topic_code47 , :v_topic_type47 , :v_survey_column47 , :v_topic_title47 , :v_topic_frame47 , :v_segment47 , :v_product_mindset47 , :v_ranking_num47 , :v_topic_code48 , :v_topic_type48 , :v_survey_column48 , :v_topic_title48 , :v_topic_frame48 , :v_segment48 , :v_product_mindset48 , :v_ranking_num48 , :v_topic_code49 , :v_topic_type49 , :v_survey_column49 , :v_topic_title49 , :v_topic_frame49 , :v_segment49 , :v_product_mindset49 , :v_ranking_num49 , :g_ranking_total, :g_ranking_total0, :g_ranking_total1, :g_ranking_total2, :g_ranking_total3, :g_ranking_total4, :g_ranking_total5, :g_ranking_total6, :g_ranking_total7, :g_ranking_total8, :g_ranking_total9, :g_ranking_total10, :g_ranking_total11, :g_ranking_total12, :g_ranking_total13, :g_ranking_total14, :g_ranking_total15, :g_ranking_total16, :g_ranking_total17, :g_ranking_total18, :g_ranking_total19, :g_ranking_total20, :g_ranking_total21, :g_ranking_total22, :g_ranking_total23, :g_ranking_total24, :g_ranking_total25, :g_ranking_total26, :g_ranking_total27, :g_ranking_total28, :g_ranking_total29, :g_ranking_total30, :g_ranking_total31, :g_ranking_total32, :g_ranking_total33, :g_ranking_total34, :g_ranking_total35, :g_ranking_total36, :g_ranking_total37, :g_ranking_total38, :g_ranking_total39, :g_ranking_total40, :g_ranking_total41, :g_ranking_total42, :g_ranking_total43, :g_ranking_total44, :g_ranking_total45, :g_ranking_total46, :g_ranking_total47, :g_ranking_total48, :g_ranking_total49,:v_ranking_total, :v_ranking_total0, :v_ranking_total1, :v_ranking_total2, :v_ranking_total3, :v_ranking_total4, :v_ranking_total5, :v_ranking_total6, :v_ranking_total7, :v_ranking_total8, :v_ranking_total9, :v_ranking_total10, :v_ranking_total11, :v_ranking_total12, :v_ranking_total13, :v_ranking_total14, :v_ranking_total15, :v_ranking_total16, :v_ranking_total17, :v_ranking_total18, :v_ranking_total19, :v_ranking_total20, :v_ranking_total21, :v_ranking_total22, :v_ranking_total23, :v_ranking_total24, :v_ranking_total25, :v_ranking_total26, :v_ranking_total27, :v_ranking_total28, :v_ranking_total29, :v_ranking_total30, :v_ranking_total31, :v_ranking_total32, :v_ranking_total33, :v_ranking_total34, :v_ranking_total35, :v_ranking_total36, :v_ranking_total37, :v_ranking_total38, :v_ranking_total39, :v_ranking_total40, :v_ranking_total41, :v_ranking_total42, :v_ranking_total43, :v_ranking_total44, :v_ranking_total45, :v_ranking_total46, :v_ranking_total47, :v_ranking_total48, :v_ranking_total49, :graph_topics, :graph_topics0, :graph_topics1, :graph_topics2, :graph_topics3, :graph_topics4, :graph_topics5, :graph_topics6, :graph_topics7, :graph_topics8, :graph_topics9, :graph_topics10, :graph_topics11, :graph_topics12, :graph_topics13, :graph_topics14, :graph_topics15, :graph_topics16, :graph_topics17, :graph_topics18, :graph_topics19, :graph_topics20, :graph_topics21, :graph_topics22, :graph_topics23, :graph_topics24, :graph_topics25, :graph_topics26, :graph_topics27, :graph_topics28, :graph_topics29, :graph_topics30, :graph_topics31, :graph_topics32, :graph_topics33, :graph_topics34, :graph_topics35, :graph_topics36, :graph_topics37, :graph_topics38, :graph_topics39, :graph_topics40, :graph_topics41, :graph_topics42, :graph_topics43, :graph_topics44, :graph_topics45, :graph_topics46, :graph_topics47, :graph_topics48, :graph_topics49, :v_mindset, :v_mindset_0,:v_mindset_1, :v_mindset_2, :v_mindset_3, :v_mindset_4, :v_mindset_5, :v_mindset_6, :v_mindset_7, :v_mindset_8,  :v_mindset_9, :v_mindset_10, :v_mindset_11, :v_mindset_12, :v_mindset_13, :v_mindset_14, :v_mindset_15, :v_mindset_16, :v_mindset_17, :v_mindset_18, :v_mindset_19, :v_mindset_20, :v_mindset_21, :v_mindset_22, :v_mindset_23, :v_mindset_24, :v_mindset_25, :v_mindset_26, :v_mindset_27, :v_mindset_28, :v_mindset_29, :v_mindset_30, :v_mindset_31, :v_mindset_32, :v_mindset_33, :v_mindset_34, :v_mindset_35, :v_mindset_36, :v_mindset_37, :v_mindset_38, :v_mindset_39, :v_mindset_40, :v_mindset_41, :v_mindset_42, :v_mindset_43, :v_mindset_44, :v_mindset_45, :v_mindset_46, :v_mindset_47, :v_mindset_48, :v_mindset_49,:mapping_file, :mapping_file_name, :commit )
    end

    def v_todo_params
        params.require(:todo).permit(:v_topic_code, :v_topic_type, :v_survey_column, :v_topic_title ,:v_topic_frame , :v_segment , :v_product_mindset , :v_ranking_num ,:v_ranking_total, :v_mindset, :v_topic_code0 , :v_topic_type0 , :v_survey_column0 , :v_topic_title0 , :v_topic_frame0 , :v_segment0 , :v_product_mindset0 , :v_ranking_num0 , :v_ranking_total0 ,:v_mindset_0, :v_topic_code1 , :v_topic_type1 , :v_survey_column1 , :v_topic_title1 , :v_topic_frame1 , :v_segment1 , :v_product_mindset1 , :v_ranking_num1 , :v_ranking_total1 ,:v_mindset_1, :v_topic_code2 , :v_topic_type2 , :v_survey_column2 , :v_topic_title2 , :v_topic_frame2 , :v_segment2 , :v_product_mindset2 , :v_ranking_num2 , :v_ranking_total2 ,:v_mindset_2, :v_topic_code3 , :v_topic_type3 , :v_survey_column3 , :v_topic_title3 , :v_topic_frame3 , :v_segment3 , :v_product_mindset3 , :v_ranking_num3 , :v_ranking_total3 ,:v_mindset_3, :v_topic_code4 , :v_topic_type4 , :v_survey_column4 , :v_topic_title4 , :v_topic_frame4 , :v_segment4 , :v_product_mindset4 , :v_ranking_num4 , :v_ranking_total4 ,:v_mindset_4, :v_topic_code5 , :v_topic_type5 , :v_survey_column5 , :v_topic_title5 , :v_topic_frame5 , :v_segment5 , :v_product_mindset5 , :v_ranking_num5 , :v_ranking_total5 ,:v_mindset_5, :v_topic_code6 , :v_topic_type6 , :v_survey_column6 , :v_topic_title6 , :v_topic_frame6 , :v_segment6 , :v_product_mindset6 , :v_ranking_num6 , :v_ranking_total6 ,:v_mindset_6, :v_topic_code7 , :v_topic_type7 , :v_survey_column7 , :v_topic_title7 , :v_topic_frame7 , :v_segment7 , :v_product_mindset7 , :v_ranking_num7 , :v_ranking_total7 ,:v_mindset_7, :v_topic_code8 , :v_topic_type8 , :v_survey_column8 , :v_topic_title8 , :v_topic_frame8 , :v_segment8 , :v_product_mindset8 , :v_ranking_num8 , :v_ranking_total8 ,:v_mindset_8, :v_topic_code9 , :v_topic_type9 , :v_survey_column9 , :v_topic_title9 , :v_topic_frame9 , :v_segment9 , :v_product_mindset9 , :v_ranking_num9 , :v_ranking_total9 ,:v_mindset_9, :v_topic_code10 , :v_topic_type10 , :v_survey_column10 , :v_topic_title10 , :v_topic_frame10 , :v_segment10 , :v_product_mindset10 , :v_ranking_num10 , :v_ranking_total10 ,:v_mindset_10, :v_topic_code11 , :v_topic_type11 , :v_survey_column11 , :v_topic_title11 , :v_topic_frame11 , :v_segment11 , :v_product_mindset11 , :v_ranking_num11 , :v_ranking_total11 ,:v_mindset_11, :v_topic_code12 , :v_topic_type12 , :v_survey_column12 , :v_topic_title12 , :v_topic_frame12 , :v_segment12 , :v_product_mindset12 , :v_ranking_num12 , :v_ranking_total12 ,:v_mindset_12, :v_topic_code13 , :v_topic_type13 , :v_survey_column13 , :v_topic_title13 , :v_topic_frame13 , :v_segment13 , :v_product_mindset13 , :v_ranking_num13 , :v_ranking_total13 ,:v_mindset_13, :v_topic_code14 , :v_topic_type14 , :v_survey_column14 , :v_topic_title14 , :v_topic_frame14 , :v_segment14 , :v_product_mindset14 , :v_ranking_num14 , :v_ranking_total14 ,:v_mindset_14, :v_topic_code15 , :v_topic_type15 , :v_survey_column15 , :v_topic_title15 , :v_topic_frame15 , :v_segment15 , :v_product_mindset15 , :v_ranking_num15 , :v_ranking_total15 ,:v_mindset_15, :v_topic_code16 , :v_topic_type16 , :v_survey_column16 , :v_topic_title16 , :v_topic_frame16 , :v_segment16 , :v_product_mindset16 , :v_ranking_num16 , :v_ranking_total16 ,:v_mindset_16, :v_topic_code17 , :v_topic_type17 , :v_survey_column17 , :v_topic_title17 , :v_topic_frame17 , :v_segment17 , :v_product_mindset17 , :v_ranking_num17 , :v_ranking_total17 ,:v_mindset_17, :v_topic_code18 , :v_topic_type18 , :v_survey_column18 , :v_topic_title18 , :v_topic_frame18 , :v_segment18 , :v_product_mindset18 , :v_ranking_num18 , :v_ranking_total18 ,:v_mindset_18, :v_topic_code19 , :v_topic_type19 , :v_survey_column19 , :v_topic_title19 , :v_topic_frame19 , :v_segment19 , :v_product_mindset19 , :v_ranking_num19 , :v_ranking_total19 ,:v_mindset_19, :v_topic_code20 , :v_topic_type20 , :v_survey_column20 , :v_topic_title20 , :v_topic_frame20 , :v_segment20 , :v_product_mindset20 , :v_ranking_num20 , :v_ranking_total20 ,:v_mindset_20, :v_topic_code21 , :v_topic_type21 , :v_survey_column21 , :v_topic_title21 , :v_topic_frame21 , :v_segment21 , :v_product_mindset21 , :v_ranking_num21 , :v_ranking_total21 ,:v_mindset_21, :v_topic_code22 , :v_topic_type22 , :v_survey_column22 , :v_topic_title22 , :v_topic_frame22 , :v_segment22 , :v_product_mindset22 , :v_ranking_num22 , :v_ranking_total22 ,:v_mindset_22, :v_topic_code23 , :v_topic_type23 , :v_survey_column23 , :v_topic_title23 , :v_topic_frame23 , :v_segment23 , :v_product_mindset23 , :v_ranking_num23 , :v_ranking_total23 ,:v_mindset_23, :v_topic_code24 , :v_topic_type24 , :v_survey_column24 , :v_topic_title24 , :v_topic_frame24 , :v_segment24 , :v_product_mindset24 , :v_ranking_num24 , :v_ranking_total24 ,:v_mindset_24, :v_topic_code25 , :v_topic_type25 , :v_survey_column25 , :v_topic_title25 , :v_topic_frame25 , :v_segment25 , :v_product_mindset25 , :v_ranking_num25 , :v_ranking_total25 ,:v_mindset_25, :v_topic_code26 , :v_topic_type26 , :v_survey_column26 , :v_topic_title26 , :v_topic_frame26 , :v_segment26 , :v_product_mindset26 , :v_ranking_num26 , :v_ranking_total26 ,:v_mindset_26, :v_topic_code27 , :v_topic_type27 , :v_survey_column27 , :v_topic_title27 , :v_topic_frame27 , :v_segment27 , :v_product_mindset27 , :v_ranking_num27 , :v_ranking_total27 ,:v_mindset_27, :v_topic_code28 , :v_topic_type28 , :v_survey_column28 , :v_topic_title28 , :v_topic_frame28 , :v_segment28 , :v_product_mindset28 , :v_ranking_num28 , :v_ranking_total28 ,:v_mindset_28, :v_topic_code29 , :v_topic_type29 , :v_survey_column29 , :v_topic_title29 , :v_topic_frame29 , :v_segment29 , :v_product_mindset29 , :v_ranking_num29 , :v_ranking_total29 ,:v_mindset_29, :v_topic_code30 , :v_topic_type30 , :v_survey_column30 , :v_topic_title30 , :v_topic_frame30 , :v_segment30 , :v_product_mindset30 , :v_ranking_num30 , :v_ranking_total30 ,:v_mindset_30, :v_topic_code31 , :v_topic_type31 , :v_survey_column31 , :v_topic_title31 , :v_topic_frame31 , :v_segment31 , :v_product_mindset31 , :v_ranking_num31 , :v_ranking_total31 ,:v_mindset_31, :v_topic_code32 , :v_topic_type32 , :v_survey_column32 , :v_topic_title32 , :v_topic_frame32 , :v_segment32 , :v_product_mindset32 , :v_ranking_num32 , :v_ranking_total32 ,:v_mindset_32, :v_topic_code33 , :v_topic_type33 , :v_survey_column33 , :v_topic_title33 , :v_topic_frame33 , :v_segment33 , :v_product_mindset33 , :v_ranking_num33 , :v_ranking_total33 ,:v_mindset_33, :v_topic_code34 , :v_topic_type34 , :v_survey_column34 , :v_topic_title34 , :v_topic_frame34 , :v_segment34 , :v_product_mindset34 , :v_ranking_num34 , :v_ranking_total34 ,:v_mindset_34, :v_topic_code35 , :v_topic_type35 , :v_survey_column35 , :v_topic_title35 , :v_topic_frame35 , :v_segment35 , :v_product_mindset35 , :v_ranking_num35 , :v_ranking_total35 ,:v_mindset_35, :v_topic_code36 , :v_topic_type36 , :v_survey_column36 , :v_topic_title36 , :v_topic_frame36 , :v_segment36 , :v_product_mindset36 , :v_ranking_num36 , :v_ranking_total36 ,:v_mindset_36, :v_topic_code37 , :v_topic_type37 , :v_survey_column37 , :v_topic_title37 , :v_topic_frame37 , :v_segment37 , :v_product_mindset37 , :v_ranking_num37 , :v_ranking_total37 ,:v_mindset_37, :v_topic_code38 , :v_topic_type38 , :v_survey_column38 , :v_topic_title38 , :v_topic_frame38 , :v_segment38 , :v_product_mindset38 , :v_ranking_num38 , :v_ranking_total38 ,:v_mindset_38, :v_topic_code39 , :v_topic_type39 , :v_survey_column39 , :v_topic_title39 , :v_topic_frame39 , :v_segment39 , :v_product_mindset39 , :v_ranking_num39 , :v_ranking_total39 ,:v_mindset_39, :v_topic_code40 , :v_topic_type40 , :v_survey_column40 , :v_topic_title40 , :v_topic_frame40 , :v_segment40 , :v_product_mindset40 , :v_ranking_num40 , :v_ranking_total40 ,:v_mindset_40, :v_topic_code41 , :v_topic_type41 , :v_survey_column41 , :v_topic_title41 , :v_topic_frame41 , :v_segment41 , :v_product_mindset41 , :v_ranking_num41 , :v_ranking_total41 ,:v_mindset_41, :v_topic_code42 , :v_topic_type42 , :v_survey_column42 , :v_topic_title42 , :v_topic_frame42 , :v_segment42 , :v_product_mindset42 , :v_ranking_num42 , :v_ranking_total42 ,:v_mindset_42, :v_topic_code43 , :v_topic_type43 , :v_survey_column43 , :v_topic_title43 , :v_topic_frame43 , :v_segment43 , :v_product_mindset43 , :v_ranking_num43 , :v_ranking_total43 ,:v_mindset_43, :v_topic_code44 , :v_topic_type44 , :v_survey_column44 , :v_topic_title44 , :v_topic_frame44 , :v_segment44 , :v_product_mindset44 , :v_ranking_num44 , :v_ranking_total44 ,:v_mindset_44, :v_topic_code45 , :v_topic_type45 , :v_survey_column45 , :v_topic_title45 , :v_topic_frame45 , :v_segment45 , :v_product_mindset45 , :v_ranking_num45 , :v_ranking_total45 ,:v_mindset_45, :v_topic_code46 , :v_topic_type46 , :v_survey_column46 , :v_topic_title46 , :v_topic_frame46 , :v_segment46 , :v_product_mindset46 , :v_ranking_num46 , :v_ranking_total46 ,:v_mindset_46, :v_topic_code47 , :v_topic_type47 , :v_survey_column47 , :v_topic_title47 , :v_topic_frame47 , :v_segment47 , :v_product_mindset47 , :v_ranking_num47 , :v_ranking_total47 ,:v_mindset_47, :v_topic_code48 , :v_topic_type48 , :v_survey_column48 , :v_topic_title48 , :v_topic_frame48 , :v_segment48 , :v_product_mindset48 , :v_ranking_num48 , :v_ranking_total48 ,:v_mindset_48, :v_topic_code49 , :v_topic_type49 , :v_survey_column49 , :v_topic_title49 , :v_topic_frame49 , :v_segment49 , :v_product_mindset49 , :v_ranking_num49 , :v_ranking_total49 ,:v_mindset_49 )
    end
    # graph params
    def g_todo_params
         params.require(:todo).permit(:g_topic_code , :g_topic_type , :graph_topics , :g_topic_title , :g_topic_frame , :g_segment , :g_product_mindset , :g_ranking_num , :g_ranking_total, :g_topic_code0 , :g_topic_type0 , :graph_topics0 , :g_topic_title0 , :g_topic_frame0 , :g_segment0 , :g_product_mindset0 , :g_ranking_num0 , :g_ranking_total0,:g_topic_code1 , :g_topic_type1 , :graph_topics1 , :g_topic_title1 , :g_topic_frame1 , :g_segment1 , :g_product_mindset1 , :g_ranking_num1 ,:g_ranking_total1,  :g_topic_code2 , :g_topic_type2 , :graph_topics2 , :g_topic_title2 , :g_topic_frame2 , :g_segment2 , :g_product_mindset2 , :g_ranking_num2 ,:g_ranking_total2, :g_topic_code3 , :g_topic_type3 , :graph_topics3 , :g_topic_title3 , :g_topic_frame3 , :g_segment3 , :g_product_mindset3 , :g_ranking_num3 ,:g_ranking_total3, :g_topic_code4 , :g_topic_type4 , :graph_topics4 , :g_topic_title4 , :g_topic_frame4 , :g_segment4 , :g_product_mindset4 , :g_ranking_num4 ,:g_ranking_total4, :g_topic_code5 , :g_topic_type5 , :graph_topics5 , :g_topic_title5 , :g_topic_frame5 , :g_segment5 , :g_product_mindset5 , :g_ranking_num5 ,:g_ranking_total5, :g_topic_code6 , :g_topic_type6 , :graph_topics6 , :g_topic_title6 , :g_topic_frame6 , :g_segment6 , :g_product_mindset6 , :g_ranking_num6 ,:g_ranking_total6, :g_topic_code7 , :g_topic_type7 , :graph_topics7 , :g_topic_title7 , :g_topic_frame7 , :g_segment7 , :g_product_mindset7 , :g_ranking_num7 ,:g_ranking_total7, :g_topic_code8 , :g_topic_type8 , :graph_topics8 , :g_topic_title8 , :g_topic_frame8 , :g_segment8 , :g_product_mindset8 , :g_ranking_num8 ,:g_ranking_total8, :g_topic_code9 , :g_topic_type9 , :graph_topics9 , :g_topic_title9 , :g_topic_frame9 , :g_segment9 , :g_product_mindset9 , :g_ranking_num9 ,:g_ranking_total9, :g_topic_code10 , :g_topic_type10 , :graph_topics10 , :g_topic_title10 , :g_topic_frame10 , :g_segment10 , :g_product_mindset10 , :g_ranking_num10 ,:g_ranking_total10, :g_topic_code11 , :g_topic_type11 , :graph_topics11 , :g_topic_title11 , :g_topic_frame11 , :g_segment11 , :g_product_mindset11 , :g_ranking_num11 ,:g_ranking_total11, :g_topic_code12 , :g_topic_type12 , :graph_topics12 , :g_topic_title12 , :g_topic_frame12 , :g_segment12 , :g_product_mindset12 , :g_ranking_num12 ,:g_ranking_total12, :g_topic_code13 , :g_topic_type13 , :graph_topics13 , :g_topic_title13 , :g_topic_frame13 , :g_segment13 , :g_product_mindset13 , :g_ranking_num13 ,:g_ranking_total13, :g_topic_code14 , :g_topic_type14 , :graph_topics14 , :g_topic_title14 , :g_topic_frame14 , :g_segment14 , :g_product_mindset14 , :g_ranking_num14 ,:g_ranking_total14, :g_topic_code15 , :g_topic_type15 , :graph_topics15 , :g_topic_title15 , :g_topic_frame15 , :g_segment15 , :g_product_mindset15 , :g_ranking_num15 ,:g_ranking_total15, :g_topic_code16 , :g_topic_type16 , :graph_topics16 , :g_topic_title16 , :g_topic_frame16 , :g_segment16 , :g_product_mindset16 , :g_ranking_num16 ,:g_ranking_total16, :g_topic_code17 , :g_topic_type17 , :graph_topics17 , :g_topic_title17 , :g_topic_frame17 , :g_segment17 , :g_product_mindset17 , :g_ranking_num17 , :g_ranking_total17,:g_topic_code18 , :g_topic_type18 , :graph_topics18 , :g_topic_title18 , :g_topic_frame18 , :g_segment18 , :g_product_mindset18 , :g_ranking_num18 ,:g_ranking_total18,  :g_topic_code19 , :g_topic_type19 , :graph_topics19 , :g_topic_title19 , :g_topic_frame19 , :g_segment19 , :g_product_mindset19 , :g_ranking_num19 ,:g_ranking_total19, :g_topic_code20 , :g_topic_type20 , :graph_topics20 , :g_topic_title20 , :g_topic_frame20 , :g_segment20 , :g_product_mindset20 , :g_ranking_num20 ,:g_ranking_total20, :g_topic_code21 , :g_topic_type21 , :graph_topics21 , :g_topic_title21 , :g_topic_frame21 , :g_segment21 , :g_product_mindset21 , :g_ranking_num21 ,:g_ranking_total21, :g_topic_code22 , :g_topic_type22 , :graph_topics22 , :g_topic_title22 , :g_topic_frame22 , :g_segment22 , :g_product_mindset22 , :g_ranking_num22 ,:g_ranking_total22, :g_topic_code23 , :g_topic_type23 , :graph_topics23 , :g_topic_title23 , :g_topic_frame23 , :g_segment23 , :g_product_mindset23 , :g_ranking_num23 ,:g_ranking_total23, :g_topic_code24 , :g_topic_type24 , :graph_topics24 , :g_topic_title24 , :g_topic_frame24 , :g_segment24 , :g_product_mindset24 , :g_ranking_num24 ,:g_ranking_total24, :g_topic_code25 , :g_topic_type25 , :graph_topics25 , :g_topic_title25 , :g_topic_frame25 , :g_segment25 , :g_product_mindset25 , :g_ranking_num25 , :g_ranking_total25,:g_topic_code26 , :g_topic_type26 , :graph_topics26 , :g_topic_title26 , :g_topic_frame26 , :g_segment26 , :g_product_mindset26 , :g_ranking_num26 ,:g_ranking_total26, :g_topic_code27 , :g_topic_type27 , :graph_topics27 , :g_topic_title27 , :g_topic_frame27 , :g_segment27 , :g_product_mindset27 , :g_ranking_num27 ,:g_ranking_total27, :g_topic_code28 , :g_topic_type28 , :graph_topics28 , :g_topic_title28 , :g_topic_frame28 , :g_segment28 , :g_product_mindset28 , :g_ranking_num28 ,:g_ranking_total28, :g_topic_code29 , :g_topic_type29 , :graph_topics29 , :g_topic_title29 , :g_topic_frame29 , :g_segment29 , :g_product_mindset29 , :g_ranking_num29 ,:g_ranking_total29, :g_topic_code30 , :g_topic_type30 , :graph_topics30 , :g_topic_title30 , :g_topic_frame30 , :g_segment30 , :g_product_mindset30 , :g_ranking_num30 , :g_ranking_total30,:g_topic_code31 , :g_topic_type31 , :graph_topics31 , :g_topic_title31 , :g_topic_frame31 , :g_segment31 , :g_product_mindset31 , :g_ranking_num31 ,:g_ranking_total31, :g_topic_code32 , :g_topic_type32 , :graph_topics32 , :g_topic_title32 , :g_topic_frame32 , :g_segment32 , :g_product_mindset32 , :g_ranking_num32 ,:g_ranking_total32, :g_topic_code33 , :g_topic_type33 , :graph_topics33 , :g_topic_title33 , :g_topic_frame33 , :g_segment33 , :g_product_mindset33 , :g_ranking_num33 , :g_ranking_total33,:g_topic_code34 , :g_topic_type34 , :graph_topics34 , :g_topic_title34 , :g_topic_frame34 , :g_segment34 , :g_product_mindset34 , :g_ranking_num34 ,:g_ranking_total34, :g_topic_code35 , :g_topic_type35 , :graph_topics35 , :g_topic_title35 , :g_topic_frame35 , :g_segment35 , :g_product_mindset35 , :g_ranking_num35 ,:g_ranking_total35, :g_topic_code36 , :g_topic_type36 , :graph_topics36 , :g_topic_title36 , :g_topic_frame36 , :g_segment36 , :g_product_mindset36 , :g_ranking_num36 ,:g_ranking_total36, :g_topic_code37 , :g_topic_type37 , :graph_topics37 , :g_topic_title37 , :g_topic_frame37 , :g_segment37 , :g_product_mindset37 , :g_ranking_num37 ,:g_ranking_total37, :g_topic_code38 , :g_topic_type38 , :graph_topics38 , :g_topic_title38 , :g_topic_frame38 , :g_segment38 , :g_product_mindset38 , :g_ranking_num38 ,:g_ranking_total38, :g_topic_code39 , :g_topic_type39 , :graph_topics39 , :g_topic_title39 , :g_topic_frame39 , :g_segment39 , :g_product_mindset39 , :g_ranking_num39 ,:g_ranking_total39, :g_topic_code40 , :g_topic_type40 , :graph_topics40 , :g_topic_title40 , :g_topic_frame40 , :g_segment40 , :g_product_mindset40 , :g_ranking_num40 ,:g_ranking_total40, :g_topic_code41 , :g_topic_type41 , :graph_topics41 , :g_topic_title41 , :g_topic_frame41 , :g_segment41 , :g_product_mindset41 , :g_ranking_num41 , :g_ranking_total41,:g_topic_code42 , :g_topic_type42 , :graph_topics42 , :g_topic_title42 , :g_topic_frame42 , :g_segment42 , :g_product_mindset42 , :g_ranking_num42 ,:g_ranking_total42, :g_topic_code43 , :g_topic_type43 , :graph_topics43 , :g_topic_title43 , :g_topic_frame43 , :g_segment43 , :g_product_mindset43 , :g_ranking_num43 ,:g_ranking_total43, :g_topic_code44 , :g_topic_type44 , :graph_topics44 , :g_topic_title44 , :g_topic_frame44 , :g_segment44 , :g_product_mindset44 , :g_ranking_num44 , :g_ranking_total44, :g_topic_code45 , :g_topic_type45 , :graph_topics45 , :g_topic_title45 , :g_topic_frame45 , :g_segment45 , :g_product_mindset45 , :g_ranking_num45 ,:g_ranking_total45, :g_topic_code46 , :g_topic_type46 , :graph_topics46 , :g_topic_title46 , :g_topic_frame46 , :g_segment46 , :g_product_mindset46 , :g_ranking_num46 ,:g_ranking_total46, :g_topic_code47 , :g_topic_type47 , :graph_topics47 , :g_topic_title47 , :g_topic_frame47 , :g_segment47 , :g_product_mindset47 , :g_ranking_num47 ,:g_ranking_total47, :g_topic_code48 , :g_topic_type48 , :graph_topics48 , :g_topic_title48 , :g_topic_frame48 , :g_segment48 , :g_product_mindset48 , :g_ranking_num48 ,:g_ranking_total48, :g_topic_code49 , :g_topic_type49 , :graph_topics49 , :g_topic_title49 , :g_topic_frame49 , :g_segment49 , :g_product_mindset49 , :g_ranking_num49, :g_ranking_total49  )
   
   end

end
