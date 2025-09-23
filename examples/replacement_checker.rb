#!/usr/bin/env ruby
# frozen_string_literal: true

# Replacement Checker Example
# Shows how to validate that all replacements were successful

require 'docx'

puts "Replacement Checker Example"
puts "=========================="

# Paths to your documents
original_template = 'template.docx'
processed_document = 'output.docx'

puts "\n📋 Files to check:"
puts "  Original: #{original_template}"
puts "  Processed: #{processed_document}"

# Step 1: Basic validation - checks for any remaining placeholders
puts "\n🔍 Step 1: Basic validation..."
validator = Docx::ReplacementValidator.validate(
  original_template,
  processed_document
)

# Show the validation report
validator.report

# Step 2: Check specific placeholders you expected to replace
puts "\n🎯 Step 2: Checking specific placeholders..."
expected_placeholders = [
  '{{ client_name }}',
  '{{ contract_date }}',
  '{{ amount }}',
  '_office_name_',
  '_partner_name_'
]

specific_results = validator.validate_expected_placeholders(expected_placeholders)

puts "\nSpecific placeholder validation:"
specific_results.each do |result|
  status = result[:successfully_replaced] ? "✅" : "❌"
  puts "#{status} #{result[:expected]}"
  
  if result[:found_in_original] > 0
    puts "    Original: #{result[:found_in_original]} instances"
    puts "    Processed: #{result[:found_in_processed]} instances"
  else
    puts "    ⚠️  Not found in original template"
  end
end

# Step 3: Final verdict
puts "\n" + "="*50
puts "FINAL VERDICT"
puts "="*50

if validator.passed?
  puts "\n🎉 SUCCESS! All placeholders were replaced correctly."
  puts "✅ Your document is ready to use."
  
  puts "\n📊 Statistics:"
  puts "  - Total placeholders found: #{validator.results[:total_placeholders_found]}"
  puts "  - Successfully replaced: #{validator.results[:successful_replacements]}"
  puts "  - Success rate: #{validator.success_rate}%"
  
else
  puts "\n❌ VALIDATION FAILED"
  puts "Some placeholders were not replaced properly."
  
  puts "\n🔧 Issues found:"
  validator.failed_placeholders.each do |failed|
    puts "  - #{failed[:placeholder]} (#{failed[:paragraph_index]})"
    puts "    Context: #{failed[:context][0..60]}..."
  end
  
  puts "\n💡 Suggestions to fix:"
  puts "  1. Use substitute_across_runs instead of substitute"
  puts "  2. Check for typos in placeholder names"
  puts "  3. Verify placeholders aren't in tables (use table processing)"
  puts "  4. Check if placeholders are fragmented by Word"
  
  # Exit with error code for CI/automation
  exit(1)
end

puts "\n🏁 Validation complete!"