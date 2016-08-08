class SlidesController < ActionController::Base

  require 'libreconv'

  def index
  end

  def se

    # Hash to be shown at the View
    @content = Hash.new(0)

    # Original filename of the uploaded pptx
    filename = params[:pptx].original_filename

    # Temporary file path of the uploaded pptx
    file_path = params[:pptx].path

    # Creating a Presentation object out of the temporary file
    presentation = Presentation.new file_path

    # Parsing has to be done BEFORE! rebuilding
    presentation.slides.each do |slide|
      # Store information inside the hash for the View
      @content.store(slide.slide_number , {slide_number: slide.slide_number, slide_title: slide.title, slide_comment: slide.comment,
                                           slide_notes: slide.notes, links: slide.links})
    end

    # Preprocessing of the pptx
    presentation.rebuild_pptx

    presentation.generate_pdf(filename)
 #   presentation.generate_images(filename)


  #  Docsplit.extract_images(file_path)
  #     Docsplit.extract_pdf(presentation.files.to_s)
  #  presentation.close
  end


  
end
