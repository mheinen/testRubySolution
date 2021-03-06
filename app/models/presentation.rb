require 'zip/filesystem'
require 'nokogiri'
class Presentation

  attr_reader :files, :narration

  def initialize(path)
    raise 'Not a valid file format.' unless (['.pptx'].include? File.extname(path).downcase)

    # Array for temporary files. Files are closed and unlinked in Presentation.close().
    @temp_files = Array.new

    # @files := Content of the uploaded pptx file.
    @files = Zip::File.open path
    narration_present?
  end

  # True if presentation contains Narration
  def narration_present?
    content_types_xml = nil
    pres_doc = @files.file.open '[Content_Types].xml' rescue nil
    content_types_xml = Nokogiri::XML::Document.parse(pres_doc) if pres_doc
    content = content_types_xml.xpath('//@Extension')
    if content.to_s.include?('m4a')
      @type = 'm4a'
      @narration = true
    end
    if content.to_s.include?('wav')
      @type = 'wav'
      @narration = true
    end
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

  # Creates a physical copy of the narration file and returns the file path in the asset pipeline
  def open_narration(slide_no)
    if @type == 'm4a'
     @files.extract("ppt/media/media#{slide_no}.m4a", "#{Rails.root}/public/audios/#{$filename}media#{slide_no}.m4a") rescue
     $filename+'media'+slide_no.to_s+'.m4a'
    end
    if @type == 'wav'

      @files.extract("ppt/media/media#{slide_no}.wav", "#{Rails.root}/public/audios/#{$filename}media#{slide_no}.wav") rescue
      $filename+'media'+slide_no.to_s+'.wav'

    end
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
          temp_file.close
          # Collect temporary files to unlink them later
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
    @destination_path
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