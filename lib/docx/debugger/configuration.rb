# frozen_string_literal: true

module Docx
  module Debugger
    class Configuration
      PLACEHOLDER_PATTERNS = {
        mustache: {
          pattern: /\{\s*(\w+)\s*\}/,
          example: '{ placeholder_name }',
          description: 'Single curly braces with spaces'
        },
        double_mustache: {
          pattern: /\{\{\s*(\w+)\s*\}\}/,
          example: '{{ placeholder_name }}',
          description: 'Double curly braces (Handlebars/Mustache style)'
        },
        underline: {
          pattern: /_(\w+)_/,
          example: '_placeholder_name_',
          description: 'Underscores on both sides'
        },
        angle: {
          pattern: /<\s*(\w+)\s*>/,
          example: '< placeholder_name >',
          description: 'Angle brackets'
        },
        square: {
          pattern: /\[\s*(\w+)\s*\]/,
          example: '[ placeholder_name ]',
          description: 'Square brackets'
        },
        dollar: {
          pattern: /\$\{(\w+)\}/,
          example: '${placeholder_name}',
          description: 'Dollar sign with curly braces (template literal style)'
        }
      }.freeze

      attr_accessor :placeholder_type, :placeholders, :custom_pattern
      attr_reader :results

      def initialize
        @placeholder_type = :double_mustache
        @placeholders = []
        @custom_pattern = nil
        @results = {}
      end

      def pattern
        return @custom_pattern if @custom_pattern
        PLACEHOLDER_PATTERNS[@placeholder_type][:pattern]
      end

      def pattern_description
        return 'Custom pattern' if @custom_pattern
        PLACEHOLDER_PATTERNS[@placeholder_type][:description]
      end

      def pattern_example
        return 'Custom regex pattern' if @custom_pattern
        PLACEHOLDER_PATTERNS[@placeholder_type][:example]
      end

      def add_placeholder(name)
        formatted = format_placeholder(name)
        @placeholders << {
          name: name,
          formatted: formatted,
          pattern: build_pattern_for(name)
        }
      end

      def set_placeholders(placeholder_list)
        @placeholders = []
        placeholder_list.each { |p| add_placeholder(p) }
      end

      private

      def format_placeholder(name)
        case @placeholder_type
        when :mustache
          "{ #{name} }"
        when :double_mustache
          "{{ #{name} }}"
        when :underline
          "_#{name}_"
        when :angle
          "< #{name} >"
        when :square
          "[ #{name} ]"
        when :dollar
          "${#{name}}"
        else
          name
        end
      end

      def build_pattern_for(name)
        case @placeholder_type
        when :mustache
          /\{\s*#{Regexp.escape(name)}\s*\}/
        when :double_mustache
          /\{\{\s*#{Regexp.escape(name)}\s*\}\}/
        when :underline
          /_#{Regexp.escape(name)}_/
        when :angle
          /<\s*#{Regexp.escape(name)}\s*>/
        when :square
          /\[\s*#{Regexp.escape(name)}\s*\]/
        when :dollar
          /\$\{#{Regexp.escape(name)}\}/
        else
          /#{Regexp.escape(name)}/
        end
      end
    end
  end
end