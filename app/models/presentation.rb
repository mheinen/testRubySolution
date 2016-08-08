class Presentation

  attr_reader :files

  def initialize(path)
    raise 'Not a valid file format.' unless (['.pptx'].include? File.extname(path).downcase)

    # Array for temporary files. Files are closed and unlinked in Presentation.close().
    @temp_files = Array.new

    # @files := Content of the uploaded pptx file.
    @files = Zip::File.open path
  end

  # Creates slides out of the corresponding files.
  def slides
    @slides = Array.new
    @files.each do |f|
      if f.name.include? 'ppt/slides/slide'
        @slides.push Slide.new(self, f.name)
      end
    end
    @slides.sort{|a,b| a.slide_num <=> b.slide_num}
  end

  # Since the slides have to be manipulated to show correctly in Libre, the pptx has to be rebuild.
  def rebuild_pptx
    Zip::File.open(@files.to_s, Zip::File::CREATE) { |zipfile|
      @slides.each do |f|
        if f.changed
          # Temporary file to store the manipulated xml
          temp_file = Tempfile.new(f.slide_file_name)
          # Store the manipulated xml inside the file
          temp_file.write(f.raw_xml)
          # Collect temporary files to close and unlink them later
          @temp_files << temp_file
          # Replace the original slide with the new one
          zipfile.replace(f.slide_xml_path, temp_file.path)
        end
      end
    }
  end

  # Generates the pdf with Libreconv
  # gem 'libreconv'
  def generate_pdf(filename)
    @destination_path = "#{Rails.root}/public/files/"+filename.remove('pptx')+'pdf'
    Libreconv.convert( @files.to_s, @destination_path, nil, 'pdf:writer_pdf_Export')
   # Docsplit.extract_pdf(@files.to_s)

    # Creates the manipulated pptx file physically
    FileUtils.mv @files.to_s, "#{Rails.root}/public/files/"+filename
  end

  # Generates Images out of the pptx into the root folder
  # Installation and dependencies at http://documentcloud.github.io/docsplit/
  def generate_images
    Docsplit.extract_images(@files.to_s)
  end

  def close
    @files.close
    # @temp_files.each do |f|
    #   f.close
    #   f.unlink
    # end
  end
end