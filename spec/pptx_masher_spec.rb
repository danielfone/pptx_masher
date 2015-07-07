require 'spec_helper'
require 'tempfile'

describe PPTXMasher do
  it 'has a version number' do
    expect(PPTXMasher::VERSION).not_to be nil
  end

  it 'does something useful' do
    output_file = Tempfile.new ['spec', '.pptx']
    output_path = output_file.path

    template = PPTXMasher::Presentation.load "spec/data/stasaph_rate_card.pptx"
    tmp_dir = template.tmp_dir

    expect(template.slides.count).to eq 1

    s = template.slides[1]
    s2 = template.add_slide s
    s2.replace_text '[TITLE]', "First update with &entities <foo>"
    s2.replace_media 'image6.jpeg', File.open('spec/data/AK00101-1.jpg')
    s3 = template.add_slide s
    s3.replace_text '[TITLE]', "Second update"

    expect(template.slides.count).to eq 3

    digest = checksum tmp_dir
    expect(digest).to eq '9567aec315db7bc69e5edc7449805d7d'

    template.save_and_close output_path

    type = `file #{output_path}`.strip
    expect(type).to include 'Zip archive data, at least v2.0 to extract'
    expect(Dir.exists? tmp_dir).to eq false

    system "open -W #{output_path}" if ENV['OPEN_PPT']
    output_file.unlink
  end

  def checksum(dir)
    files = Dir["#{dir}/**/*"].reject{|f| File.directory?(f)}
    content = files.map{|f| File.read(f)}.join
    require 'digest/md5'
    Digest::MD5.hexdigest content
  end

end
