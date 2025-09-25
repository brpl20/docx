require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {}, doc = nil)
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @document = doc
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = ''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['color'] = "##{font_color}" if font_color
          styles['text-align'] = alignment if alignment
          html_tag(:p, content: html, styles: styles)
        end


        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_attribute = @node.at_xpath('w:pPr//w:sz//@w:val')

          return @font_size unless size_attribute

          size_attribute.value.to_i / 2
        end

        def font_color
          color_tag = @node.xpath('w:r//w:rPr//w:color').first
          color_tag ? color_tag.attributes['val'].value : nil
        end

        def style
          return nil unless @document

          @document.style_name_of(style_id) ||
            @document.default_paragraph_style
        end

        def style_id
          style_property.get_attribute('w:val')
        end

        def style=(identifier)
          id = @document.styles_configuration.style_of(identifier).id

          style_property.set_attribute('w:val', id)
        end

        alias_method :style_id=, :style=
        alias_method :text, :to_s

        # Consolidates adjacent text runs with identical formatting
        # This helps solve the issue where placeholders get fragmented across multiple runs
        def consolidate_text_runs
          runs = text_runs
          return if runs.empty?

          consolidated = []
          current_group = [runs.first]
          current_formatting = runs.first.formatting

          runs[1..-1].each do |run|
            if run.formatting == current_formatting
              # Same formatting, add to current group
              current_group << run
            else
              # Different formatting, consolidate current group and start new one
              consolidate_group(current_group) if current_group.length > 1
              consolidated << current_group
              current_group = [run]
              current_formatting = run.formatting
            end
          end

          # Don't forget the last group
          consolidate_group(current_group) if current_group.length > 1
          consolidated << current_group
        end

        # Performs text substitution across all text runs in the paragraph
        # This solves the issue where placeholders are split across multiple runs
        def substitute_across_runs(pattern, replacement)
          # First, get all text nodes in order
          all_text_nodes = []
          text_runs.each do |run|
            run.instance_variable_get(:@text_nodes).each do |text_node|
              all_text_nodes << text_node
            end
          end

          return if all_text_nodes.empty?

          # Concatenate all text content
          full_text = all_text_nodes.map(&:content).join('')
          
          # Check if pattern exists in the full text
          return unless full_text.match?(pattern)

          # Perform the replacement
          new_text = full_text.gsub(pattern, replacement)

          # Redistribute the new text back to the nodes
          # Try to maintain original text node boundaries where possible
          if all_text_nodes.length == 1
            # Simple case: single text node
            all_text_nodes.first.content = new_text
          else
            # Complex case: multiple text nodes
            # We'll put all the text in the first node and clear the others
            # This maintains formatting while ensuring substitution works
            all_text_nodes.first.content = new_text
            all_text_nodes[1..-1].each { |node| node.content = '' }
          end

          # Update the text in text runs
          text_runs.each { |run| run.send(:reset_text) }
        end

        # Performs block-based text substitution across all text runs
        def substitute_across_runs_with_block(pattern, &block)
          all_text_nodes = []
          text_runs.each do |run|
            run.instance_variable_get(:@text_nodes).each do |text_node|
              all_text_nodes << text_node
            end
          end

          return if all_text_nodes.empty?

          full_text = all_text_nodes.map(&:content).join('')
          return unless full_text.match?(pattern)

          new_text = full_text.gsub(pattern) { |_matched| 
            block.call(Regexp.last_match)
          }

          if all_text_nodes.length == 1
            all_text_nodes.first.content = new_text
          else
            all_text_nodes.first.content = new_text
            all_text_nodes[1..-1].each { |node| node.content = '' }
          end

          text_runs.each { |run| run.send(:reset_text) }
        end

        # Performs block-based text substitution with enhanced regex patterns
        # Uses more powerful regex patterns to avoid placeholder overlapping
        # If pattern is a string, it will be wrapped with word boundary protection: /(?<![_\w])pattern(?![_\w])/
        def substitute_across_runs_with_block_regex(pattern, &block)
          # Convert string patterns to regex with word boundary protection
          if pattern.is_a?(String)
            pattern = /(?<![_\w])#{Regexp.escape(pattern)}(?![_\w])/
          end
          all_text_nodes = []
          text_runs.each do |run|
            run.instance_variable_get(:@text_nodes).each do |text_node|
              all_text_nodes << text_node
            end
          end

          return if all_text_nodes.empty?

          full_text = all_text_nodes.map(&:content).join('')
          return unless full_text.match?(pattern)

          new_text = full_text.gsub(pattern) { |_matched| 
            block.call(Regexp.last_match)
          }

          if all_text_nodes.length == 1
            all_text_nodes.first.content = new_text
          else
            all_text_nodes.first.content = new_text
            all_text_nodes[1..-1].each { |node| node.content = '' }
          end

          text_runs.each { |run| run.send(:reset_text) }
        end

        private

        # Helper method to consolidate a group of text runs with same formatting
        def consolidate_group(group)
          return if group.length <= 1

          # Merge all text into the first run
          first_run = group.first
          combined_text = group.map(&:text).join('')
          
          # Set the combined text in the first run
          first_run.text = combined_text

          # Remove the other runs from the XML
          group[1..-1].each do |run|
            run.node.remove
          end
        end

        def style_property
          properties&.at_xpath('w:pStyle') || properties&.add_child('<w:pStyle/>').first
        end

        # Returns the alignment if any, or nil if left
        def alignment
          @node.at_xpath('.//w:jc/@w:val')&.value
        end
      end
    end
  end
end
