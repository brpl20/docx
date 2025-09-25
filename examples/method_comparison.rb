#!/usr/bin/env ruby
# frozen_string_literal: true

# Comparison: Old vs New substitution methods
# Shows the evolution from basic substitution to advanced regex-protected substitution

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'

puts "=" * 80
puts "METHOD COMPARISON: DOCX PLACEHOLDER SUBSTITUTION EVOLUTION"
puts "=" * 80

puts "\n📚 AVAILABLE METHODS:"
puts "-" * 50
puts "1. substitute_across_runs(pattern, replacement)"
puts "   • Basic text replacement"
puts "   • Static replacement only"
puts ""
puts "2. substitute_across_runs_with_block(pattern, &block)"
puts "   • Dynamic replacement with block logic"
puts "   • Manual regex patterns required for safety"
puts ""
puts "3. substitute_across_runs_with_block_regex(pattern, &block) ⭐ NEW!"
puts "   • Automatic word boundary protection"
puts "   • Clean string-based patterns"
puts "   • Advanced regex support when needed"

puts "\n" + "=" * 80
puts "EXAMPLE 1: BASIC REPLACEMENT"
puts "=" * 80

puts "\n❌ OLD METHOD 1 - Basic (Not recommended for complex cases):"
puts 'paragraph.substitute_across_runs(/_company_name_/, "ACME Corp")'
puts "• Risk: May replace partial matches like 'my_company_name_here'"
puts "• No dynamic logic"

puts "\n✅ NEW METHOD - With automatic protection:"
puts 'paragraph.substitute_across_runs_with_block_regex("_company_name_") do |match|'
puts '  "ACME Corp"'
puts 'end'
puts "• Safe: Only replaces exact '_company_name_' with word boundaries"
puts "• Supports dynamic logic"

puts "\n" + "=" * 80
puts "EXAMPLE 2: DYNAMIC REPLACEMENT WITH SAFETY"
puts "=" * 80

puts "\n❌ OLD METHOD 2 - Manual regex (Error-prone):"
puts 'paragraph.substitute_across_runs_with_block(/(?<![_\\w])_partner_name_(?![_\\w])/) do |match|'
puts '  partner_data["name"]'
puts 'end'
puts "• Verbose regex syntax"
puts "• Easy to make mistakes with escaping"
puts "• Hard to read and maintain"

puts "\n✅ NEW METHOD - Clean and safe:"
puts 'paragraph.substitute_across_runs_with_block_regex("_partner_name_") do |match|'
puts '  partner_data["name"]'
puts 'end'
puts "• Clean, readable syntax"
puts "• Automatic word boundary protection"
puts "• Less error-prone"

puts "\n" + "=" * 80
puts "EXAMPLE 3: COMPLEX PATTERNS"
puts "=" * 80

puts "\n⚡ ADVANCED - Custom regex when needed:"
puts 'pattern = /(?<![_\\w])_office_(\\w+)_(?![_\\w])/'
puts 'paragraph.substitute_across_runs_with_block_regex(pattern) do |match|'
puts '  field = match[1]  # captured group'
puts '  office_data[field]'
puts 'end'
puts "• Full regex power when needed"
puts "• Captures groups available"
puts "• Best of both worlds"

puts "\n" + "=" * 80
puts "REAL-WORLD EXAMPLE: LEGAL DOCUMENT PROCESSING"
puts "=" * 80

# Sample data structure
document_data = {
  'society' => {
    'name' => 'Innovation Partners LLC',
    'city' => 'New York',
    'state' => 'NY',
    'pro_labore' => true
  },
  'partners' => [
    { 'name' => 'John', 'last_name' => 'Smith', 'percentage' => 60 },
    { 'name' => 'Jane', 'last_name' => 'Doe', 'percentage' => 40 }
  ]
}

puts "\n🏢 Processing legal document for:"
puts "   Company: #{document_data['society']['name']}"
puts "   Partners: #{document_data['partners'].length}"

puts "\n📝 Replacement patterns using NEW METHOD:"
puts "-" * 50

# Simulate document processing
replacements = [
  {
    pattern: "_society_name_",
    value: document_data['society']['name'],
    description: "Company name"
  },
  {
    pattern: "_society_location_",
    value: "#{document_data['society']['city']}, #{document_data['society']['state']}",
    description: "Company location"
  },
  {
    pattern: "_partner_count_",
    value: document_data['partners'].length.to_s,
    description: "Number of partners"
  },
  {
    pattern: "_total_percentage_",
    value: "#{document_data['partners'].sum { |p| p['percentage'] }}%",
    description: "Total ownership percentage"
  }
]

replacements.each_with_index do |replacement, index|
  puts "#{index + 1}. Pattern: '#{replacement[:pattern]}'"
  puts "   Value: '#{replacement[:value]}'"
  puts "   Description: #{replacement[:description]}"
  puts "   Code:"
  puts "   paragraph.substitute_across_runs_with_block_regex(\"#{replacement[:pattern]}\") do |match|"
  puts "     \"#{replacement[:value]}\""
  puts "   end"
  puts
end

puts "\n🔄 Conditional replacement example:"
puts "-" * 50
puts "# Pro labore clause (conditional)"
puts 'paragraph.substitute_across_runs_with_block_regex("_pro_labore_clause_") do |match|'
puts "  if document_data['society']['pro_labore']"
puts '    "Pro labore payments will be distributed monthly."'
puts "  else"
puts '    ""  # Remove clause if not applicable'
puts "  end"
puts "end"

pro_labore_result = document_data['society']['pro_labore'] ? 
  "Pro labore payments will be distributed monthly." : "[REMOVED]"
puts "Result: #{pro_labore_result}"

puts "\n" + "=" * 80
puts "MIGRATION GUIDE: OLD → NEW"
puts "=" * 80

migration_examples = [
  {
    old: 'substitute_across_runs(/_name_/, "John")',
    new: 'substitute_across_runs_with_block_regex("_name_") { "John" }'
  },
  {
    old: 'substitute_across_runs_with_block(/(?<![_\\w])_field_(?![_\\w])/) { value }',
    new: 'substitute_across_runs_with_block_regex("_field_") { value }'
  },
  {
    old: 'substitute_across_runs_with_block(/(?<![_\\w])_%_#{num}_(?![_\\w])/) { "#{percent}%" }',
    new: 'substitute_across_runs_with_block_regex("_%_#{num}_") { "#{percent}%" }'
  }
]

migration_examples.each_with_index do |example, index|
  puts "\n#{index + 1}. Migration Example:"
  puts "   ❌ OLD: #{example[:old]}"
  puts "   ✅ NEW: #{example[:new]}"
end

puts "\n" + "=" * 80
puts "SUMMARY: WHY USE THE NEW METHOD?"
puts "=" * 80
puts "✅ Safer - Automatic word boundary protection"
puts "✅ Cleaner - No manual regex for common cases"  
puts "✅ Flexible - Still supports custom regex when needed"
puts "✅ Maintainable - Easier to read and modify"
puts "✅ Robust - Handles DOCX text run fragmentation"
puts "✅ Future-proof - Built for complex document processing"

puts "\n🚀 Ready to upgrade your DOCX processing!"
puts "=" * 80