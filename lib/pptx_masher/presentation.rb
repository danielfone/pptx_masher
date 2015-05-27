require 'pptx_masher/quick_zip'
require 'pptx_masher/slide'

module PPTXMasher
  class Presentation
    attr_reader :tmp_dir

    def self.load(src)
      dest = File.join Dir.tmpdir, "pptx_masher", SecureRandom.uuid
      QuickZip.extract src, dest
      new dest
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
      # Copy media?
      new_slide.insert_content_type
      new_slide.insert_rels
      new_slide.insert_meta
      new_slide
    end

    def save(path)
      QuickZip.compress tmp_dir, path
    end

    def slide_count
      Dir.glob(tmp_dir+"/ppt/slides/slide*.xml").count
    end

    def media_count
      Dir.glob(tmp_dir+"/ppt/media/*").count
    end

  end
end
