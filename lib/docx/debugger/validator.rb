# frozen_string_literal: true

require 'securerandom'

module Docx
  module Debugger
    class Validator
      def initialize(document, config)
        @document = document
        @config = config
      end

      def validate_placeholder(placeholder_info)
        result = {
          placeholder: placeholder_info,
          success: false,
          found_count: 0,
          paragraph_indices: [],
          replaced_count: 0,
          error: nil
        }

        begin
          # First, search for the placeholder in all paragraphs
          search_results = search_placeholder(placeholder_info)
          result[:found_count] = search_results[:count]
          result[:paragraph_indices] = search_results[:indices]

          if result[:found_count] == 0
            # Try to find if it's fragmented
            fragmented_results = search_fragmented(placeholder_info)
            if fragmented_results[:found]
              result[:error] = "Placeholder found but fragmented across text runs. Use substitute_across_runs method."
              result[:found_count] = fragmented_results[:count]
              result[:paragraph_indices] = fragmented_results[:indices]
            else
              result[:error] = "Placeholder not found in document"
            end
          else
            # Test replacement
            test_result = test_replacement(placeholder_info)
            if test_result[:success]
              result[:success] = true
              result[:replaced_count] = test_result[:replaced_count]
            else
              result[:error] = test_result[:error]
            end
          end
        rescue => e
          result[:error] = "Validation error: #{e.message}"
        end

        result
      end

      private

      def search_placeholder(placeholder_info)
        count = 0
        indices = []
        
        @document.paragraphs.each_with_index do |paragraph, index|
          text = paragraph.text
          if text.match?(placeholder_info[:pattern])
            count += text.scan(placeholder_info[:pattern]).length
            indices << index + 1  # 1-indexed for user display
          end
        end

        { count: count, indices: indices }
      end

      def search_fragmented(placeholder_info)
        # Check if parts of the placeholder exist but are fragmented
        name = placeholder_info[:name]
        found = false
        count = 0
        indices = []

        @document.paragraphs.each_with_index do |paragraph, index|
          # Get all text run contents separately
          run_texts = paragraph.text_runs.map(&:text)
          
          # Check if the name appears somewhere in the paragraph
          full_text = paragraph.text
          if full_text.include?(name) || contains_parts_of_placeholder?(run_texts, placeholder_info[:formatted])
            found = true
            count += 1
            indices << index + 1
          end
        end

        { found: found, count: count, indices: indices }
      end

      def contains_parts_of_placeholder?(run_texts, formatted_placeholder)
        # Check if the placeholder is split across runs
        # For example, "{{name}}" might be split as ["{{", "na", "me", "}}"]
        joined = run_texts.join('')
        
        # Remove the formatted placeholder's special characters and check
        cleaned_placeholder = formatted_placeholder.gsub(/[{}\[\]<>_\s$]/, '')
        joined.include?(cleaned_placeholder)
      end

      def test_replacement(placeholder_info)
        begin
          test_value = "REPLACED_#{SecureRandom.hex(4)}"
          replaced_count = 0
          
          # Create a temporary copy to test
          temp_paragraphs = @document.paragraphs.dup
          
          temp_paragraphs.each do |paragraph|
            original_text = paragraph.text
            
            # Try both methods
            # Method 1: Traditional run-by-run
            paragraph.each_text_run do |run|
              if run.text.match?(placeholder_info[:pattern])
                run.substitute(placeholder_info[:pattern], test_value)
                replaced_count += 1
              end
            end
            
            # Method 2: Across runs (if method 1 didn't work)
            if original_text == paragraph.text && original_text.match?(placeholder_info[:pattern])
              paragraph.substitute_across_runs(placeholder_info[:pattern], test_value)
              if paragraph.text.include?(test_value)
                replaced_count += 1
              end
            end
          end
          
          if replaced_count > 0
            { success: true, replaced_count: replaced_count }
          else
            { success: false, error: "Replacement failed - placeholder might be malformed" }
          end
        rescue => e
          { success: false, error: "Replacement test failed: #{e.message}" }
        end
      end
    end
  end
end