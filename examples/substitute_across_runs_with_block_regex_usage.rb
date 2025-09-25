#!/usr/bin/env ruby
# frozen_string_literal: true

# Example: Using substitute_across_runs_with_block_regex method
# This method provides automatic word boundary protection for string patterns
# and advanced regex substitution capabilities

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'

puts "=" * 70
puts "SUBSTITUTE_ACROSS_RUNS_WITH_BLOCK_REGEX EXAMPLES"
puts "=" * 70

# Example 1: Simple string pattern with automatic word boundary protection
puts "\n1. Simple String Pattern (Automatic Word Boundaries)"
puts "-" * 50

# Create a sample document or open an existing one
# doc = Docx::Document.open('template.docx')

# Mock paragraph for demonstration
class MockParagraph
  def substitute_across_runs_with_block_regex(pattern, &block)
    # Simulate the method behavior
    if pattern.is_a?(String)
      regex_pattern = /(?<![_\w])#{Regexp.escape(pattern)}(?![_\w])/
      puts "  String pattern '#{pattern}' converted to: #{regex_pattern.inspect}"
    else
      puts "  Custom regex pattern: #{pattern.inspect}"
    end
    
    # Simulate replacement
    yield(nil) if block_given?
  end
end

paragraph = MockParagraph.new

# Simple usage with string - automatic word boundary protection
paragraph.substitute_across_runs_with_block_regex("_company_name_") do |match|
  result = "ACME Corporation Ltd."
  puts "  ✅ Replaced _company_name_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_date_") do |match|
  result = "December 25, 2024"
  puts "  ✅ Replaced _date_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_%_") do |match|
  result = "85%"
  puts "  ✅ Replaced _%_ → #{result}"
  result
end

puts "\n2. Dynamic String Patterns"
puts "-" * 50

# Dynamic patterns with variables
partner_number = 2
paragraph.substitute_across_runs_with_block_regex("_partner_name_#{partner_number}_") do |match|
  result = "John Smith"
  puts "  ✅ Replaced _partner_name_#{partner_number}_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_%_#{partner_number}_") do |match|
  result = "15%"
  puts "  ✅ Replaced _%_#{partner_number}_ → #{result}"
  result
end

puts "\n3. Custom Complex Regex Patterns"
puts "-" * 50

# Custom regex for more advanced matching
paragraph.substitute_across_runs_with_block_regex(/(?<![_\w])_office_\w+_(?![_\w])/) do |match|
  # This would match _office_name_, _office_address_, _office_phone_, etc.
  result = "Various office information"
  puts "  ✅ Replaced office pattern → #{result}"
  result
end

# Pattern with capture groups
paragraph.substitute_across_runs_with_block_regex(/(?<![_\w])_(\w+)_value_(?![_\w])/) do |match|
  field_name = $1 # Captured group
  result = "Value for #{field_name}"
  puts "  ✅ Replaced _#{field_name}_value_ → #{result}"
  result
end

puts "\n4. Real-World Usage Example"
puts "-" * 50

# Simulate a real document processing scenario
data = {
  'company' => {
    'name' => 'Tech Solutions Inc.',
    'address' => '123 Innovation Street',
    'city' => 'San Francisco',
    'state' => 'CA'
  },
  'contract' => {
    'date' => '2024-12-25',
    'value' => 50000,
    'percentage' => 75
  }
}

puts "Processing document with data:"
puts "  Company: #{data['company']['name']}"
puts "  Location: #{data['company']['city']}, #{data['company']['state']}"
puts "  Contract Value: $#{data['contract']['value']}"
puts

# Process placeholders
paragraph.substitute_across_runs_with_block_regex("_company_name_") do |match|
  result = data['company']['name']
  puts "  ✅ _company_name_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_company_address_") do |match|
  result = data['company']['address']
  puts "  ✅ _company_address_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_company_location_") do |match|
  result = "#{data['company']['city']}, #{data['company']['state']}"
  puts "  ✅ _company_location_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_contract_value_") do |match|
  result = "$#{data['contract']['value'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, ',')}"
  puts "  ✅ _contract_value_ → #{result}"
  result
end

paragraph.substitute_across_runs_with_block_regex("_%_") do |match|
  result = "#{data['contract']['percentage']}%"
  puts "  ✅ _%_ → #{result}"
  result
end

puts "\n5. Conditional Replacements"
puts "-" * 50

# Conditional logic based on data
show_optional_clause = true

paragraph.substitute_across_runs_with_block_regex("_optional_clause_") do |match|
  if show_optional_clause
    result = "This optional clause is included in the contract."
    puts "  ✅ _optional_clause_ → #{result[0..40]}..."
  else
    result = ""
    puts "  ✅ _optional_clause_ → [REMOVED]"
  end
  result
end

puts "\n6. Error Handling and Fallbacks"
puts "-" * 50

# Safe replacement with fallbacks
paragraph.substitute_across_runs_with_block_regex("_missing_field_") do |match|
  begin
    # Simulate accessing a missing field
    result = data['missing']['field'] || "N/A"
  rescue
    result = "[DATA NOT AVAILABLE]"
    puts "  ⚠️  _missing_field_ → #{result} (fallback used)"
  end
  result
end

puts "\n" + "=" * 70
puts "KEY BENEFITS OF substitute_across_runs_with_block_regex:"
puts "=" * 70
puts "✅ Automatic word boundary protection for string patterns"
puts "✅ Prevents placeholder overlapping issues"
puts "✅ Cleaner, more readable code"
puts "✅ Still supports custom regex when needed"
puts "✅ Handles text runs fragmentation automatically"
puts "✅ Perfect for DOCX template processing"

puts "\n" + "=" * 70
puts "USAGE PATTERNS:"
puts "=" * 70
puts '• Simple: paragraph.substitute_across_runs_with_block_regex("_placeholder_")'
puts '• Dynamic: paragraph.substitute_across_runs_with_block_regex("_field_#{number}_")'
puts '• Custom regex: paragraph.substitute_across_runs_with_block_regex(/pattern/)'

puts "\n" + "=" * 70