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
          zip_file.add(zip_path, path)
        end
      end
    end

  end
end
