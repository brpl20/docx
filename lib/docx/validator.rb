# frozen_string_literal: true

module Docx
  class ReplacementValidator
    attr_reader :original_document, :processed_document, :results

    def initialize(original_path, processed_path)
      @original_path = original_path
      @processed_path = processed_path
      @original_document = Docx::Document.open(original_path)
      @processed_document = Docx::Document.open(processed_path)
      @results = {
        total_placeholders_found: 0,
        successful_replacements: 0,
        failed_replacements: 0,
        missed_placeholders: [],
        replacement_details: [],
        validation_passed: false
      }
    end

    def self.validate(original_path, processed_path, expected_patterns = [])
      validator = new(original_path, processed_path)
      validator.validate_replacements(expected_patterns)
      validator
    end

    def validate_replacements(expected_patterns = [])
      # Default patterns to check for common placeholder formats
      default_patterns = [
        /_\w+_/,                    # Underline: _placeholder_
        /\{\s*\w+\s*\}/,           # Mustache: { placeholder }
        /\{\{\s*\w+\s*\}\}/,       # Double mustache: {{ placeholder }}
        /<\s*\w+\s*>/,             # Angle: < placeholder >
        /\[\s*\w+\s*\]/,           # Square: [ placeholder ]
        /\$\{\w+\}/                # Dollar: ${placeholder}
      ]

      patterns_to_check = expected_patterns.empty? ? default_patterns : expected_patterns

      # Check original document for placeholders
      original_placeholders = find_placeholders(@original_document, patterns_to_check)
      processed_placeholders = find_placeholders(@processed_document, patterns_to_check)

      @results[:total_placeholders_found] = original_placeholders.length

      # Compare original vs processed
      original_placeholders.each do |placeholder_info|
        placeholder = placeholder_info[:placeholder]
        
        # Check if this placeholder still exists in processed document
        still_exists = processed_placeholders.any? { |p| p[:placeholder] == placeholder }
        
        if still_exists
          @results[:failed_replacements] += 1
          @results[:missed_placeholders] << placeholder_info
        else
          @results[:successful_replacements] += 1
        end

        @results[:replacement_details] << {
          placeholder: placeholder,
          original_paragraph: placeholder_info[:paragraph_index],
          original_text: placeholder_info[:context],
          replaced: !still_exists,
          pattern_matched: placeholder_info[:pattern_matched]
        }
      end

      @results[:validation_passed] = @results[:failed_replacements] == 0

      self
    end

    def validate_expected_placeholders(expected_placeholders)
      # Validate that specific placeholders were replaced
      validation_results = []

      expected_placeholders.each do |expected|
        pattern = case expected
                  when String
                    # Convert string to appropriate pattern
                    if expected.start_with?('_') && expected.end_with?('_')
                      /#{Regexp.escape(expected)}/
                    elsif expected.match?(/\{\s*\w+\s*\}/)
                      /\{\s*#{Regexp.escape(expected.gsub(/[{}]/, '').strip)}\s*\}/
                    else
                      /#{Regexp.escape(expected)}/
                    end
                  when Regexp
                    expected
                  else
                    raise ArgumentError, "Expected String or Regexp, got #{expected.class}"
                  end

        found_in_original = find_pattern_in_document(@original_document, pattern)
        found_in_processed = find_pattern_in_document(@processed_document, pattern)

        validation_results << {
          expected: expected.to_s,
          found_in_original: found_in_original.length,
          found_in_processed: found_in_processed.length,
          successfully_replaced: found_in_original.length > 0 && found_in_processed.length == 0,
          locations_original: found_in_original,
          locations_processed: found_in_processed
        }
      end

      validation_results
    end

    def report
      puts "\n" + "="*70
      puts "REPLACEMENT VALIDATION REPORT"
      puts "="*70
      
      puts "\n📊 Summary:"
      puts "  Total placeholders found: #{@results[:total_placeholders_found]}"
      puts "  ✅ Successfully replaced: #{@results[:successful_replacements]}"
      puts "  ❌ Failed to replace: #{@results[:failed_replacements]}"
      puts "  🎯 Success rate: #{success_rate}%"

      if @results[:validation_passed]
        puts "\n🎉 VALIDATION PASSED - All placeholders were replaced!"
      else
        puts "\n⚠️ VALIDATION FAILED - Some placeholders were not replaced!"
        
        puts "\n❌ Missed placeholders:"
        @results[:missed_placeholders].each do |missed|
          puts "  - #{missed[:placeholder]} (Paragraph #{missed[:paragraph_index]})"
          puts "    Context: ...#{missed[:context][0..60]}..."
          puts "    Pattern: #{missed[:pattern_matched]}"
        end

        puts "\n💡 Suggestions:"
        puts "  1. Check if placeholders are fragmented across XML nodes"
        puts "  2. Use substitute_across_runs instead of substitute"
        puts "  3. Verify placeholder spelling and format"
        puts "  4. Check if placeholders are in tables or special sections"
      end

      puts "\n📋 Detailed Results:"
      @results[:replacement_details].each_with_index do |detail, i|
        status = detail[:replaced] ? "✅" : "❌"
        puts "  #{i+1}. #{status} #{detail[:placeholder]}"
        puts "     Location: Paragraph #{detail[:original_paragraph]}"
        puts "     Pattern: #{detail[:pattern_matched]}"
      end

      puts "\n" + "="*70
    end

    def success_rate
      return 0 if @results[:total_placeholders_found] == 0
      ((@results[:successful_replacements].to_f / @results[:total_placeholders_found]) * 100).round(2)
    end

    def passed?
      @results[:validation_passed]
    end

    def failed_placeholders
      @results[:missed_placeholders]
    end

    private

    def find_placeholders(document, patterns)
      placeholders = []
      
      document.paragraphs.each_with_index do |paragraph, para_index|
        text = paragraph.text
        
        patterns.each do |pattern|
          text.scan(pattern) do |match|
            matched_text = $&  # The full match
            placeholders << {
              placeholder: matched_text,
              paragraph_index: para_index + 1,
              context: text,
              pattern_matched: pattern
            }
          end
        end
      end

      # Also check tables
      document.tables.each_with_index do |table, table_index|
        table.rows.each_with_index do |row, row_index|
          row.cells.each_with_index do |cell, cell_index|
            cell.paragraphs.each_with_index do |paragraph, para_index|
              text = paragraph.text
              
              patterns.each do |pattern|
                text.scan(pattern) do |match|
                  matched_text = $&
                  placeholders << {
                    placeholder: matched_text,
                    paragraph_index: "Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}, Para #{para_index + 1}",
                    context: text,
                    pattern_matched: pattern
                  }
                end
              end
            end
          end
        end
      end

      placeholders.uniq { |p| [p[:placeholder], p[:paragraph_index]] }
    end

    def find_pattern_in_document(document, pattern)
      locations = []
      
      document.paragraphs.each_with_index do |paragraph, para_index|
        text = paragraph.text
        if text.match?(pattern)
          matches = text.scan(pattern)
          matches.each do |match|
            locations << {
              paragraph_index: para_index + 1,
              matched_text: match.is_a?(Array) ? match.join : match,
              context: text[0..100]
            }
          end
        end
      end

      locations
    end
  end
end