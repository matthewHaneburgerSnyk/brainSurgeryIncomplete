require 'rubygems'
require 'fileutils'
require 'zip'
require 'rake'



#file = "NUEDEXTA-2test.zip"
#path ="./public/uploads/#{file}/" #this should work in the app but not out of it
#ext = ""



#FileUtils.cd("/Users/rl250177/Desktop/rubyBrain/graphing_tool/public/uploads/NUEDEXTA-2test.xlsx/")



#puts "file name before #{file}"


#ext = "zip"


class ZipFileGenerator
    
    # Initialize with the directory to zip and the location of the output archive.
    def initialize(inputDir, outputFile)
        @inputDir = inputDir
        @outputFile = outputFile
    end
    
    # Zip the input directory.
    def write()
        entries = Dir.entries(@inputDir); entries.delete("."); entries.delete("..")
        io = Zip::File.open(@outputFile, Zip::File::CREATE);
        
        writeEntries(entries, "", io)
        io.close();
    end
    
    # A helper method to make the recursion work.
    private
    def writeEntries(entries, path, io)
        
        entries.each { |e|
            zipFilePath = path == "" ? e : File.join(path, e)
            diskFilePath = File.join(@inputDir, zipFilePath)
            puts "Deflating " + diskFilePath
            if  File.directory?(diskFilePath)
                io.mkdir(zipFilePath)
                subdir =Dir.entries(diskFilePath); subdir.delete("."); subdir.delete("..")
                writeEntries(subdir, zipFilePath, io)
                else
                io.get_output_stream(zipFilePath) { |f| f.puts(File.open(diskFilePath, "rb").read())}
            end
        }
    end
    
end
#graphs_OCALIVAPatientDataNEWTOOLTEST_01.xlsx
