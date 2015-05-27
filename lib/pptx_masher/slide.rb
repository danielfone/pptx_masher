require 'fileutils'

module PPTXMasher
  class Slide
    attr_reader :presentation, :number

    def initialize(presentation, number)
      @presentation = presentation
      @number = Integer(number)
    end

    def slide_xml_path
      @slide_xml_path ||= "/ppt/slides/slide#{@number}.xml"
    end

    def slide_xml_full_path
      @slide_xml_full_path ||= "#{presentation_path}#{slide_xml_path}"
    end

    def rels_path
      @rels_path ||= "#{presentation_path}/ppt/slides/_rels/slide#{number}.xml.rels"
    end

    def notes_xml_path
      @notes_xml_path ||= "/ppt/notesSlides/notesSlide#{number}.xml"
    end

    def notes_xml_full_path
      @notes_xml_full_path ||= "#{presentation_path}#{notes_xml_path}"
    end

    def notes_rels_path
      @notes_rels_path ||= "#{presentation_path}/ppt/notesSlides/_rels/notesSlide#{number}.xml.rels"
    end

    def update_rels
      gsub_file rels_path, /notesSlide\d+\.xml/, "notesSlide#{number}.xml"
      gsub_file notes_rels_path, /slide\d+\.xml/, "slide#{number}.xml"
    end

    def replace_text(pattern, replacement)
      gsub_file slide_xml_full_path, pattern, replacement
    end

    def replace_media(dest, src)
      extension = dest.split('.').last
      new_name = "image#{presentation.media_count+1}.#{extension}"
      gsub_file rels_path, dest, new_name
      FileUtils.copy src, "#{presentation_path}/ppt/media/#{new_name}"
    end

    def insert_content_type
      append_xml "#{presentation_path}/[Content_Types].xml", 'Types',
        %Q[<Override PartName="#{slide_xml_path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>]
      append_xml "#{presentation_path}/[Content_Types].xml", 'Types',
        %Q[<Override PartName="#{notes_xml_path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>]
    end

    def insert_rels
      append_xml "#{presentation_path}/ppt/_rels/presentation.xml.rels", 'Relationships',
        %Q[<Relationship Id="#{relationship_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide#{number}.xml"/>]
    end

    def insert_meta
      append_xml "#{presentation_path}/ppt/presentation.xml", 'p:sldIdLst',
        %Q[<p:sldId id="#{id}" r:id="#{relationship_id}"/>]
    end

    def relationship_id
      @relationship_id ||= "rId#{number+1000}"
    end

    def id
      @id ||= "#{number+256}"
    end

  private

    def append_xml(file, node, xml)
      closing_tag = "</#{node}>"
      xml += closing_tag
      gsub_file file, closing_tag, xml
    end

    def gsub_file(path, pattern, replacement)
      text = File.read(path)
      text.gsub! pattern, replacement
      File.open(path, "w") {|f| f.print text }
    end

    def presentation_path
      @presentation_path ||= presentation.tmp_dir
    end

  end
end
