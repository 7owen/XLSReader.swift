Pod::Spec.new do |s|
  s.name         = "XLSReader.swift"
  s.version      = "0.0.1"
  s.summary      = "A Swift Framework that can read .xls Files."
  s.description  = <<-DESC
                    A Swift Framework that can read .xls Files.
                   DESC
  s.homepage     = "https://github.com/7owen/XLSReader.swift"
  s.license      = { :type => "MIT", :file => "LICENSE" }
  s.author       = { "7owen" => "a@lgw.im" }
  s.platform     = :ios
  s.platform     = :ios, "8.0"
  s.source       = { :git => "https://github.com/7owen/XLSReader.swift.git", :branch => "develop" }

  s.source_files = "libxls/**/*.{h,c}"
  #s.exclude_files= "ofoPaySDK/ofoPaySDK/SwiftyRSA"
  s.requires_arc = true
  s.libraries = "iconv"
  #s.dependency "SwiftyJSON", '~> 4.0.0'
  #s.dependency "Moya", '~> 11.0.2'
  #s.dependency "Alamofire", '~> 4.7.1'

  # s.subspec "libxls" do |ss|
  #   ss.source_files = "libxls/**/*.{h,c}"
  # end

end
