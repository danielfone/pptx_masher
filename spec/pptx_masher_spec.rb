require 'spec_helper'

describe PPTXMasher do
  it 'has a version number' do
    expect(PPTXMasher::VERSION).not_to be nil
  end

  it 'does something useful' do
    output_path = "/Users/danielfone/Desktop/cojoined.pptx"
    File.delete output_path if File.exists? output_path
    template = PPTXMasher::Presentation.load "/Users/danielfone/Desktop/stasaph_rate_card.pptx"
    s = template.slides[1]
    s2 = template.add_slide s
    s2.replace_text '[TITLE]', "First update"
    s2.replace_media 'image6.jpeg', '/Users/danielfone/Desktop/AK00101-1.jpg'
    s3 = template.add_slide s
    s3.replace_text '[TITLE]', "Second update"
    template.save output_path
  end
end
