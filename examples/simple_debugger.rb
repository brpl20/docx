#!/usr/bin/env ruby
# frozen_string_literal: true

# Simple Debugger Example
# This example shows the simplest way to debug placeholders in a document

require 'docx'

puts "Simple Debugger Example"
puts "======================"

# Method 1: Super simple - one line check
puts "\n1. Quick check (one line):"
Docx::Debugger.quick_check(
  'template.docx',
  :double_mustache,
  ['name', 'date', 'amount']
)

# Method 2: Basic configuration
puts "\n2. Basic configuration:"
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :underline
  config.set_placeholders(['office_name', 'partner_name'])
end

results = debugger.debug!

# Check results
if results[:failed] == 0
  puts "\n✅ All placeholders found and validated!"
  puts "📄 A replacer class has been generated for you."
else
  puts "\n❌ Some placeholders need attention."
  puts "Check the output above for details."
end

puts "\nDone! 🎉"