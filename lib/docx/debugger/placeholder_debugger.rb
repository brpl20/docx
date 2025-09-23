# frozen_string_literal: true

require 'securerandom'
require_relative 'configuration'
require_relative 'validator'
require_relative 'base_replacer_generator'

module Docx
  module Debugger
    class PlaceholderDebugger
      attr_reader :config, :document, :results

      def initialize(template_path)
        @config = Configuration.new
        @template_path = template_path
        @document = Docx::Document.open(template_path)
        @results = {
          total: 0,
          successful: 0,
          failed: 0,
          details: []
        }
      end

      def configure
        yield(@config) if block_given?
        self
      end

      def debug!
        puts "\n" + "="*60
        puts "DOCX PLACEHOLDER DEBUGGER"
        puts "="*60
        puts "\nTemplate: #{@template_path}"
        puts "Pattern Type: #{@config.placeholder_type}"
        puts "Pattern: #{@config.pattern_example}"
        puts "-"*60

        validate_placeholders
        display_results
        generate_replacer_if_successful

        @results
      end

      def validate_placeholders
        validator = Validator.new(@document, @config)
        
        @config.placeholders.each do |placeholder_info|
          result = validator.validate_placeholder(placeholder_info)
          @results[:total] += 1
          
          if result[:success]
            @results[:successful] += 1
          else
            @results[:failed] += 1
          end
          
          @results[:details] << result
        end
      end

      def display_results
        puts "\n📊 VALIDATION RESULTS:"
        puts "-"*60
        
        @results[:details].each_with_index do |detail, index|
          status_icon = detail[:success] ? "✅" : "❌"
          puts "\n#{index + 1}. #{status_icon} #{detail[:placeholder][:formatted]}"
          
          if detail[:success]
            puts "   Found: #{detail[:found_count]} occurrence(s)"
            puts "   Paragraphs: #{detail[:paragraph_indices].join(', ')}"
            puts "   Test replacement: SUCCESS"
          else
            puts "   Error: #{detail[:error]}"
            if detail[:found_count] == 0
              puts "   Suggestion: Check if placeholder exists in document"
              puts "              or if it's fragmented across XML nodes"
            end
          end
        end

        puts "\n" + "="*60
        puts "SUMMARY:"
        puts "  Total placeholders: #{@results[:total]}"
        puts "  ✅ Successful: #{@results[:successful]}"
        puts "  ❌ Failed: #{@results[:failed]}"
        puts "  Success Rate: #{success_rate}%"
        puts "="*60
      end

      def generate_replacer_if_successful
        if @results[:failed] == 0 && @results[:total] > 0
          puts "\n🎉 All placeholders validated successfully!"
          puts "Generating base replacer class..."
          
          generator = BaseReplacerGenerator.new(@config, @results)
          file_path = generator.generate!
          
          puts "✅ Base replacer class generated: #{file_path}"
          puts "\nYou can now use this class to replace placeholders in your documents."
        elsif @results[:failed] > 0
          puts "\n⚠️  Cannot generate base replacer class due to validation failures."
          puts "Please fix the issues above and run the debugger again."
        end
      end

      def test_replacement(save_as: nil)
        temp_doc = Docx::Document.open(@template_path)
        replacements = {}
        
        @config.placeholders.each do |placeholder|
          random_value = "TEST_#{SecureRandom.hex(4)}"
          replacements[placeholder[:name]] = random_value
          
          temp_doc.paragraphs.each do |paragraph|
            paragraph.substitute_across_runs(placeholder[:pattern], random_value)
          end
        end
        
        if save_as
          temp_doc.save(save_as)
          puts "\n📄 Test document saved: #{save_as}"
        end
        
        replacements
      end

      private

      def success_rate
        return 0 if @results[:total] == 0
        ((@results[:successful].to_f / @results[:total]) * 100).round(2)
      end
    end
  end
end