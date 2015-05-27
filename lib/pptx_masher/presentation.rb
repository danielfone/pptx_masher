require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'
require 'securerandom'

module PPTXMasher
  class Presentation
    attr_reader :tmp_dir

    def self.load(path)
      warn "Extract into separate class"
      out_path = File.join Dir.tmpdir, "pptx_masher", SecureRandom.uuid
      Zip::File.open(path) do |zip_file|
        zip_file.each do |f|
          f_path = File.join(out_path, f.name)
          FileUtils.mkdir_p(File.dirname(f_path))
          zip_file.extract(f, f_path) unless File.exist?(f_path)
        end
      end
      new out_path
    end

    def initialize(path)
      @tmp_dir = path
    end

    def slides
      @slides ||= Hash.new { |h, i| h[i] = Slide.new(self, i) }
    end

    def add_slide(slide)
      n = slide_count + 1
      new_slide = slides[n]
      FileUtils.copy slide.rels_path, new_slide.rels_path
      FileUtils.copy slide.slide_xml_full_path, new_slide.slide_xml_full_path
      # Copy media
      new_slide.insert_content_type
      new_slide.insert_rels
      new_slide.insert_meta
      new_slide
    end

    def save(path)
      warn "Extract into separate class"
      Zip::File.open(path, Zip::File::CREATE) do |zip_file|
        Dir.glob("#{tmp_dir}/**/*", ::File::FNM_DOTMATCH).each do |path|
          zip_path = path.gsub("#{tmp_dir}/","")
          next if zip_path == "." || zip_path == ".." || zip_path.match(/DS_Store/)
          begin
            zip_file.add(zip_path, path)
          rescue Zip::ZipEntryExistsError
            raise "#{path} already exists!"
          end
        end
      end
      true
    end

    def slide_count
      Dir.glob(tmp_dir+"/ppt/slides/slide*.xml").count
    end

  end

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

    def replace_text(pattern, replacement)
      gsub_file slide_xml_full_path, pattern, replacement
    end

    def replace_media(dest, src)
      extension = dest.split('.').last
      new_name = "#{SecureRandom.uuid}.#{extension}"
      gsub_file rels_path, dest, new_name
      FileUtils.copy src, "#{presentation_path}/ppt/media/#{new_name}"
    end

    def insert_content_type
      append_xml "#{presentation_path}/[Content_Types].xml", 'Types',
        %Q[<Override PartName="#{slide_xml_path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>]
    end

    def insert_rels
      append_xml "#{presentation_path}/ppt/_rels/presentation.xml.rels", 'Relationships',
        %Q[<Relationship Id="#{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide#{number}.xml"/>]
    end

    def insert_meta
      append_xml "#{presentation_path}/ppt/presentation.xml", 'p:sldIdLst',
        %Q[<p:sldId id="257" r:id="#{id}"/>]
    end

    def id
      @id ||= "rId#{number+1000}"
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
      File.open(path, "w") {|f| f.puts text }
    end

    def presentation_path
      @presentation_path ||= presentation.tmp_dir
    end

  end

end
