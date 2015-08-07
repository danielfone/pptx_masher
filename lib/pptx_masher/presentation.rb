require 'tmpdir'
require 'fileutils'
require 'pptx_masher/quick_zip'
require 'pptx_masher/slide'

module PPTXMasher
  class Presentation
    attr_reader :tmp_dir

    def self.load(src)
      dest = Dir.mktmpdir "pptx_masher"
      QuickZip.extract src, dest
      new dest
    end

    def initialize(path)
      @tmp_dir = path
    end

    def slides
      @slides ||= Hash[
        1.upto(slide_count).map {|i| [i, Slide.new(self, i)] }
      ]
    end

    def add_slide(slide)
      n = slide_count + 1
      new_slide = slides[n] = Slide.new(self, n)
      safe_copy slide.rels_path,           new_slide.rels_path
      safe_copy slide.slide_xml_full_path, new_slide.slide_xml_full_path
      safe_copy slide.notes_rels_path,     new_slide.notes_rels_path
      safe_copy slide.notes_xml_full_path, new_slide.notes_xml_full_path
      # Copy media?
      new_slide.update_rels
      new_slide.insert_content_type
      new_slide.insert_rels
      new_slide.insert_meta
      new_slide
    end

    def save(path)
      QuickZip.compress tmp_dir, path
    end

    def close
      FileUtils.remove_entry_secure tmp_dir
    end

    def safe_copy(src, dest)
      FileUtils.mkdir_p File.dirname dest
      FileUtils.cp src, dest
    end

    def save_and_close(path)
      save path and close
    end

    def slide_count
      Dir.glob(tmp_dir+"/ppt/slides/slide*.xml").count
    end

    def media_count
      Dir.glob(tmp_dir+"/ppt/media/*").count
    end

  end
end
