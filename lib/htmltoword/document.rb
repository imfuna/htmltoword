module Htmltoword
  class Document
    include XSLTHelper

    class << self
      include TemplatesHelper
      def create(content, template_name = nil, extras = false, page_setup = {})
        template_name += extension if template_name && !template_name.end_with?(extension)
        document = new(template_file(template_name))
        document.replace_files(content, extras)
        document.generate(page_setup)
      end

      def create_and_save(content, file_path, template_name = nil, extras = false, page_setup = {})
        File.open(file_path, 'wb') do |out|
          out << create(content, template_name, extras, page_setup)
        end
      end

      def create_with_content(template, content, extras = false, page_setup = {})
        template += extension unless template.end_with?(extension)
        document = new(template_file(template))
        document.replace_files(content, extras)
        document.generate(page_setup)
      end

      def extension
        '.docx'
      end

      def doc_xml_file
        'word/document.xml'
      end

      def numbering_xml_file
        'word/numbering.xml'
      end

      def header_xml_file
        'word/header1.xml'
      end

      def footer_xml_file
        'word/footer1.xml'
      end

      def relations_xml_file
        'word/_rels/document.xml.rels'
      end
      # this one isn't in the template file, so we have to create it from scratch
      def header_relations_xml_file
        'word/_rels/header1.xml.rels'
      end

      # this one isn't in the template file, so we have to create it from scratch
      def footer_relations_xml_file
        'word/_rels/footer1.xml.rels'
      end

      def content_types_xml_file
        '[Content_Types].xml'
      end
    end

    def initialize(template_path)
      @replaceable_files = {}
      @template_path = template_path
      @image_files = []
    end

    #
    # Generate a string representing the contents of a docx file.
    #
    def generate(page_setup = {})
      Zip::File.open(@template_path) do |template_zip|
        buffer = Zip::OutputStream.write_buffer do |out|
          template_zip.each do |entry|
            out.put_next_entry entry.name
            if @replaceable_files[entry.name] && entry.name == Document.doc_xml_file
              source = entry.get_input_stream.read
              # the source has only 1 sectPr and we change the doc dimensions based on the page setup
              source = replace_dimensions(source, page_setup)
              # Change only the body of document. TODO: Improve this...
              source = source.sub(/(<w:body>)((.|\n)*?)(<w:sectPr)/, "\\1#{@replaceable_files[entry.name]}\\4")
              out.write(source)
            elsif  @replaceable_files[entry.name] && entry.name == Document.header_xml_file
              source = entry.get_input_stream.read
              # Change only the body of document. TODO: Improve this...

              Rails.logger.info("header before:#{source}")
              source = source.sub(/<w:tbl>.*<\/w:tbl>/, "#{@replaceable_files[entry.name]}")
              Rails.logger.info("header after:#{source}")

              out.write(source)
            # elsif  @replaceable_files[entry.name] && entry.name == Document.footer_xml_file
            #   source = entry.get_input_stream.read
            #   Change only the body of document. TODO: Improve this...
              # Rails.logger.info("footer before:#{source}")
              # source = source.sub(/<w:tbl>.*<\/w:tbl>/, "#{@replaceable_files[entry.name]}")
              # Rails.logger.info("footer after:#{source}")
              # out.write(source)

            elsif @replaceable_files[entry.name]
              out.write(@replaceable_files[entry.name])
              Rails.logger.info("footer: #{entry.get_input_stream.read}")  if  entry.name == Document.footer_xml_file
            elsif entry.name == Document.content_types_xml_file
              raw_file = entry.get_input_stream.read
              content_types = @image_files.empty? ? raw_file : inject_image_content_types(raw_file)

              out.write(content_types)
            else
              out.write(template_zip.read(entry.name))
            end
          end
          unless @image_files.empty?
          #stream the image files into the media folder using open-uri
            @image_files.each do |hash|
              out.put_next_entry("word/media/#{hash[:filename]}")
              open(hash[:url], 'rb') do |f|
                out.write(f.read)
              end
            end
          end
        #   add the header rels as its not already in the template docx
          out.put_next_entry(Document.header_relations_xml_file)
          out.write(@replaceable_files[Document.header_relations_xml_file])
        #   add the footer rels as its not already in the template docx
          out.put_next_entry(Document.footer_relations_xml_file)
          out.write(@replaceable_files[Document.footer_relations_xml_file])
        end
        buffer.string
      end
    end

    def replace_files(html, extras = false)
      html = '<body></body>' if html.nil? || html.empty?
      original_source = Nokogiri::HTML(html.gsub(/>\s+</, '><'))
      source = xslt(stylesheet_name: 'cleanup').transform(original_source)
      Rails.logger.info("1:#{source}")
      transform_and_replace(source, xslt_path('numbering'), Document.numbering_xml_file)
      transform_and_replace(source, xslt_path('relations'), Document.relations_xml_file)
      # add in header and footer file
      transform_and_replace(source, xslt_path('header'), Document.header_xml_file)
      transform_and_replace(source, xslt_path('footer'), Document.footer_xml_file)
      transform_and_replace(source, xslt_path('footer_relations'), Document.footer_relations_xml_file)
      transform_doc_xml(source, extras)
      local_images(source)
      output_header_relations
    end

    def transform_doc_xml(source, extras = false)
      transformed_source = xslt(stylesheet_name: 'cleanup').transform(source)
      Rails.logger.info("4:#{transformed_source}")
      transformed_source = xslt(stylesheet_name: 'inline_elements').transform(transformed_source)
      Rails.logger.info("5:#{transformed_source}")
      transform_and_replace(transformed_source, document_xslt(extras), Document.doc_xml_file, extras)
    end

    private

    def replace_dimensions(source, page_setup)
      str = case
              when page_setup[:size] == 'A4' && page_setup[:orientation] == 'landscape'
                '<w:pgSz w:w="16839" w:h="11907" w:orient="landscape"/>'
              when page_setup[:size] == 'A4' && page_setup[:orientation] == 'portrait'
                '<w:pgSz w:w="11907" w:h="16839" />'
              when page_setup[:size] == 'USLetter' && page_setup[:orientation] == 'landscape'
                '<w:pgSz w:w="15842" w:h="12242" w:orient="landscape"/>'
              when page_setup[:size] == 'USLetter' && page_setup[:orientation] == 'portrait'
                '<w:pgSz w:w="12240" w:h="15840"/>'
              else
                '<w:pgSz w:w="11907" w:h="16839" />'
            end
      source.gsub(/<w:pgSz(.|\n)*?\/>/, str)
    end

    def transform_and_replace(source, stylesheet_path, file, remove_ns = false)
      stylesheet = xslt(stylesheet_path: stylesheet_path)
      content = stylesheet.apply_to(source)
      content.gsub!(/\s*xmlns:(\w+)="(.*?)\s*"/, '') if remove_ns
      Rails.logger.info("6:#{content}")
      @replaceable_files[file] = content
    end

    #generates an array of hashes with filename and full url
    #for all images to be embeded in the word document
    def local_images(source)
      source.css('img').each_with_index do |image,i|
        filename = image['data-filename'] ? image['data-filename'] : image['src'].split("/").last
        # remove the ? from the image name
        ext = File.extname(filename).delete(".").split('?').first.downcase

        @image_files << { filename: "image#{i+1}.#{ext}", url: image['src'], ext: ext }
      end
    end
    #  we have one image in the header and this is the first in the document and so we can hard code this file
    #  logo can be png or jpg
    def output_header_relations
      ext =   @image_files.first[:ext]
      @replaceable_files[Document.header_relations_xml_file] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.' +
          ext +
          '"/></Relationships>'
    end

    #get extension from filename and clean to match content_types
    def content_type_from_extension(ext)
      ext == "jpg" ? "jpeg" : ext
    end

    #inject the required content_types into the [content_types].xml file...
    def inject_image_content_types(source)
      doc = Nokogiri::XML(source)

      #get a list of all extensions currently in content_types file
      existing_exts = doc.css("Default").map { |node| node.attribute("Extension").value }.compact

      #get a list of extensions we need for our images
      required_exts = @image_files.map{ |i| i[:ext] }

      #workout which required extensions are missing from the content_types file
      missing_exts = (required_exts - existing_exts).uniq

      #inject missing extensions into document
      missing_exts.each do |ext|
        doc.at_css("Types").add_child( "<Default Extension='#{ext}' ContentType='image/#{content_type_from_extension(ext)}'/>")
      end

      #return the amended source to be saved into the zip
      doc.to_s
    end
  end
end
