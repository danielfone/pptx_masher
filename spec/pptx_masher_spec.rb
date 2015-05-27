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

    s = template.slides[1]
    s2 = template.add_slide s
    s2.replace_text '[TITLE]', "First update"
    s2.replace_media 'image6.jpeg', 'spec/data/AK00101-1.jpg'
    s3 = template.add_slide s
    s3.replace_text '[TITLE]', "Second update"

    digest = checksum tmp_dir
    expect(digest).to eq '4eb11a3f93f263f23f19d2d50dac334a'

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
