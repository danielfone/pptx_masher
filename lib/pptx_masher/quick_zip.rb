require 'zip/filesystem'
require 'fileutils'

module PPTXMasher
  module QuickZip
    module_function

    def extract(src, dest)
      Zip::File.open src do |zip_file|
        zip_file.each do |src_file|
          dest_file = File.join dest, src_file.name
          FileUtils.mkdir_p File.dirname dest_file
          zip_file.extract src_file, dest_file
        end
      end
    end

    def compress(src, dest)
      src_files = Dir.glob "#{src}/**/*", ::File::FNM_DOTMATCH

      Zip::File.open dest, Zip::File::CREATE do |zip_file|
        src_files.each do |path|
          zip_path = path.gsub "#{src}/", ''
          next if zip_path == "." || zip_path == ".."
          make_zip_path(zip_file, path, zip_path)
        end
      end
    end

    def make_zip_path(zip_file, path, zip_path)
      if File.directory?(path)
        zip_file.mkdir(zip_path)
      else
        zip_file.get_output_stream zip_path do |io|
          IO.copy_stream path, io
        end
      end
    end

  end
end
