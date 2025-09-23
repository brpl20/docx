#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('lib', __dir__))
require 'docx'
require 'docx/debugger'

# List of placeholders to check
placeholders = %w[
  office_name
  partner_qualification
  office_name
  office_city
  office_state
  office_address
  office_zip_code
  office_total_value
  office_quotes
  office_quote_value
  partner_full_name
  partner_total_quotes
  office_quote_value
  partner_sum
  partner_full_name
]

unique_placeholders = placeholders.uniq

# Run the debugger
debugger = Docx::Debugger.analyze('tests/CS-TEMPLATE.docx') do |config|
  config.placeholder_type = :underline
  config.set_placeholders(unique_placeholders)
end

results = debugger.debug!

# Test replacement if all placeholders are valid
if results[:failed] == 0
  puts "\n" + '=' * 70
  puts 'TESTING REPLACEMENT WITH SAMPLE VALUES'
  puts '=' * 70

  test_replacements = debugger.test_replacement(save_as: 'tests/CS-TEMPLATE-test-output.docx')

  puts "\n✅ Test document created: tests/CS-TEMPLATE-test-output.docx"
  puts "\nSample replacements used:"
  test_replacements.each do |key, value|
    puts "  #{key}: #{value}"
  end

  puts "\n" + '=' * 70
  puts 'GENERATED REPLACER CLASS'
  puts '=' * 70
  puts "\nThe replacer class has been generated. You can now use it like this:\n\n"

  puts <<~RUBY
    require_relative 'underline_replacer'

    replacer = UnderlineReplacer.new('tests/CS-TEMPLATE.docx')

    # Set your actual values
    replacer.office_name = "Smith & Associates Law Firm"
    replacer.partner_qualification = "Senior Partner"
    replacer.office_city = "New York"
    replacer.office_state = "NY"
    replacer.office_address = "123 Main Street, Suite 500"
    replacer.office_zip_code = "10001"
    replacer.office_total_value = "$1,500,000"
    replacer.office_quotes = "150"
    replacer.office_quote_value = "$10,000"
    replacer.partner_full_name = "John Smith"
    replacer.parner_total_quotes = "75"  # Note: typo preserved from original
    replacer.partner_sum = "$750,000"
    replacer.percentage = "50%"
    replacer.total_quotes = "150"
    replacer.sum_percentage = "100%"

    # Generate the final document
    replacer.process!('tests/CS-FINAL.docx')
  RUBY
else
  puts "\n" + '=' * 70
  puts '⚠️  ISSUES FOUND - DETAILED ANALYSIS'
  puts '=' * 70

  puts "\nFailed placeholders that need attention:"
  results[:details].each do |detail|
    next if detail[:success]

    puts "\n❌ #{detail[:placeholder][:formatted]}"
    puts "   Error: #{detail[:error]}"

    # Provide suggestions based on error type
    if detail[:error].include?('not found')
      puts '   💡 Suggestion: Check if the placeholder exists in the document'
      puts "                 or if there's a typo in the placeholder name"
    elsif detail[:error].include?('fragmented')
      puts '   💡 Suggestion: This placeholder is split across multiple XML nodes'
      puts '                 The generated replacer will handle this automatically'
    end
  end

  # Check for possible typos
  puts "\n" + '=' * 70
  puts '📝 POSSIBLE TYPOS DETECTED'
  puts '=' * 70

  if unique_placeholders.include?('parner_total_quotes')
    puts "\n⚠️  'parner_total_quotes' might be a typo"
    puts "   Did you mean 'partner_total_quotes'?"
  end

  # List successful placeholders for reference
  successful = results[:details].select { |d| d[:success] }
  if successful.any?
    puts "\n" + '=' * 70
    puts '✅ SUCCESSFULLY VALIDATED PLACEHOLDERS'
    puts '=' * 70
    successful.each do |detail|
      puts "\n✅ #{detail[:placeholder][:formatted]}"
      puts "   Found: #{detail[:found_count]} occurrence(s)"
      puts "   Located in paragraphs: #{detail[:paragraph_indices].join(', ')}"
    end
  end
end

puts "\n" + '=' * 70
puts 'DEBUGGING COMPLETE'
puts '=' * 70
