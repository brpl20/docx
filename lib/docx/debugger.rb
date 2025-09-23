# frozen_string_literal: true

require_relative 'debugger/configuration'
require_relative 'debugger/placeholder_debugger'
require_relative 'debugger/validator'
require_relative 'debugger/base_replacer_generator'

module Docx
  module Debugger
    class << self
      def analyze(template_path, &block)
        debugger = PlaceholderDebugger.new(template_path)
        debugger.configure(&block) if block_given?
        debugger
      end

      def quick_check(template_path, placeholder_type, placeholders)
        debugger = PlaceholderDebugger.new(template_path)
        debugger.configure do |config|
          config.placeholder_type = placeholder_type
          config.set_placeholders(placeholders)
        end
        debugger.debug!
      end
    end
  end
end