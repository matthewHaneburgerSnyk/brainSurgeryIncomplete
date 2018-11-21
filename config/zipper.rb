require 'rubygems'
require 'fileutils'
require 'zip'
require 'rake'



#file = "NUEDEXTA-2test.xlsx"
#FileUtils.cd("/Users/rl250177/Desktop/rubyBrain/graphing_tool/public/uploads/NUEDEXTA-2test.xlsx/")
#need to specify a working directory




class Zipper
   def initialize(file)
   
   @file = file
   @ext = ""
   
   
   #
   # Changes file extension to .zip
   # Have to pass it a file name and be in the right directory
   #
   def convert_to_zip(file)
       
       a = file.split(".")

       file_path = a[0] + ".xlsx"
       
       FileUtils.cd("./public/uploads/#{file_path}/")
       FileUtils.mv( file, file.ext("zip"))
       @ext = "zip"
       
   end
   
   
   #
   # Changes file extension to xlsx
   # Have to pass it a file name and be in the right directory
   #
   def convert_to_xlsx(file)
       
       a = file.split(".")
       
       file_path = a[0] + ".xlsx"
       
    
       FileUtils.cd("./public/uploads/#{file}/")
       FileUtils.mv( file, file.ext("zip"))
       @ext = "zip"
   end
   
   #
   # Extracts all files in the zip to folder 'zip_contents'
   # Have to pass it a file name and be in the right directory
   #
   def extract_files(file)
       
       Zip::File.open(file) do |zip_file|
           zip_file.each { |f|
               f_path=File.join("zip_contents", f.name)
               FileUtils.mkdir_p(File.dirname(f_path))
               zip_file.extract(f, f_path) unless File.exist?(f_path) }
       end
       
   end
   
   
   #
   # Compresses all files in the directory
   #
   def compress_files(file)
   
   
     
       
       
   end
   
   
   
end






#Dir.glob('sample_mapping.xlsx').each do |filename|
#    FileUtils.mv f, "#{File.dirname(f)}/#{File.basename(f,'.*')}.zip"
#end

#Zip::File.open('foo.zip') do |zip_file|
#    # Handle entries one by one
#    zip_file.each do |entry|
#        # Extract to file/directory/symlink
#       puts "Extracting #{entry.name}"
#        entry.extract(dest_file)
#
#       # Read into memory
#        content = entry.get_input_stream.read
#    end
#
#    # Find specific entry
#    entry = zip_file.glob('*.csv').first
#    puts entry.get_input_stream.read
#end
