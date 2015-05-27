require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module PPTXMasher
  class Presentation
    attr_reader :tmp_dir

    def self.load(path)
      warn "Extract into separate class"
      out_path = File.join Dir.tmpdir, "pptx_masher", "extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"
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

  end

end
