class Slide

  attr_reader :presentation,
              :slide_number,
              :slide_xml_path,
              :slide_file_name,
              :relation_xml,
              :slide_xml,
              :changed,
              :slide_notes_xml


  def initialize(presentation, slide_xml_path)
    @presentation = presentation
    @slide_xml_path = slide_xml_path
    @slide_number = extract_slide_number_from_path slide_xml_path
    @slide_notes_xml_path = "ppt/notesSlides/notesSlide#{@slide_number}.xml"
    @slide_file_name = extract_slide_file_name_from_path slide_xml_path
    @changed = false
    parse_slide
    parse_slide_notes
    parse_relation
    fix_libre

  end

  # Preprocessing of the xml to make sure, that the pdf looks nice when generated with Libre
  def fix_libre
    # Searches for mc:Choice nodes. These contain things that can not be shown correctly in Libre, instead Libre will show the remaining mc:Fallback nodes
    i = 0
    alternate_content = @slide_xml.xpath('//mc:AlternateContent', 'mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    while i < alternate_content.length
      choice = @slide_xml.search('//mc:Choice', 'mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006')
      choice.remove
      i += 1
      @changed = true
    end

    # Sometimes tables inside presentations contain misplaced images, these are removed
    i = 0
    table_pic = @slide_xml.xpath('//a:tcPr/a:blipFill')
    while i < table_pic.length
      blip = @slide_xml.xpath('//a:tcPr/a:blipFill/a:blip')
      blip.remove
      i += 1
      @changed = true
    end
  end


  # Parse every link on a slide and return them as an array
  def links
    # Since a slide can have multiple links, we need an array
    @links = Hash.new
    # a:r elements containing a link
    links_xml = @slide_xml.xpath('//a:r[a:rPr/a:hlinkClick]')
    if links_xml.empty?
    else
      links_xml.each do |f|
        # Text that is shown on the slide
        shown_text = f.xpath('a:t/text()')
        # Gather the ID of a link to search at the relation file
        link_id = f.xpath('string(a:rPr/a:hlinkClick/@r:id)')
        # Sadly the relation xml file is not well formed, we need to manipulate a string
        xml_string = @relation_xml.xpath('//@Id | //@Target')
        # Build an array of he single ids
        id_array = xml_string.to_s.scan(/rId./)
        # Build an array of the Target value
        content_array = xml_string.to_s.split(/rId./)
        # Remove the first empty entry
        content_array.delete_at(0)
        # Push an array into the array for the view
        if @links.has_key?(content_array[id_array.index(link_id)])
          @links[content_array[id_array.index(link_id)]] = @links[content_array[id_array.index(link_id)]] + ' ' + shown_text.to_s
        else
          @links.store(content_array[id_array.index(link_id)], shown_text.to_s)
        end
    #    @links.push([shown_text.to_s, content_array[id_array.index(link_id)]])
      end
    end
    @links
  end

  def raw_xml
    @slide_xml
  end

  def parse_slide
    slide_doc = @presentation.files.file.open @slide_xml_path
    @slide_xml = Nokogiri::XML::Document.parse slide_doc
  end

  def parse_slide_notes
    slide_notes_doc = @presentation.files.file.open @slide_notes_xml_path rescue nil
    @slide_notes_xml = Nokogiri::XML::Document.parse(slide_notes_doc) if slide_notes_doc
  end

  def parse_relation
    @relation_xml_path = "ppt/slides/_rels/#{@slide_file_name}.rels"
    if @presentation.files.file.exist? @relation_xml_path
      relation_doc = @presentation.files.file.open @relation_xml_path
      @relation_xml = Nokogiri::XML::Document.parse relation_doc
    end
  end

  def content
    content_elements @slide_xml
  end

  # Checks if there are any comments on the slide. If yes, it starts the parsing process
  def comment
    if @relation_xml.xpath('//@Target[starts-with(. , "../comments/comment")]').empty?
    else
       parse_comment(@relation_xml.xpath('//@Target[starts-with(. , "../comments/comment")]'))
    end
  end

  # Returns plain notes text
  def notes
    if slide_notes_xml.nil?
    else
    @slide_notes_xml.search('//a:t/text()')
    end
  end

  # Returns plain comment text
  def parse_comment(path)
    comment_path = path.to_s.remove('..')
    @comment_xml_path = 'ppt' + comment_path

    comment_file = @presentation.files.file.open @comment_xml_path
    @comment_xml = Nokogiri::XML::Document.parse comment_file

    @comment = @comment_xml.xpath('//p:text/text()')
  end

  def notes_content
    content_elements @slide_notes_xml
  end

  def title
    title_elements = title_elements(@slide_xml)
    title_elements.join(" ") if title_elements.length > 0
  end

  def slide_num
    @slide_xml_path.match(/slide([0-9]*)\.xml$/)[1].to_i
  end

  private

  def extract_slide_number_from_path path
    path.gsub('ppt/slides/slide', '').gsub('.xml', '').to_i
  end

  def extract_slide_file_name_from_path path
    path.gsub('ppt/slides/', '')
  end

  def title_elements(xml)
    shape_elements(xml).select{ |shape| element_is_title(shape) }
  end

  def content_elements(xml)
    xml.xpath('//a:t').collect{ |node| node.text }
  end

  def shape_elements(xml)
    xml.xpath('//p:sp')
  end

  def element_is_title(shape)
    shape.xpath('.//p:nvSpPr/p:nvPr/p:ph').select{ |prop| prop['type'] == 'title' || prop['type'] == 'ctrTitle' }.length > 0
  end

end