#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'

puts "\n" + "="*70
puts "VALIDATED INTELLIGENT REPLACEMENT WITH ERROR CHECKING"
puts "="*70

# Load the JSON data
data = JSON.parse(File.read('tests/test_information.json'))

# Define expected placeholders that should be replaced
expected_underline_placeholders = [
  '_office_name_',
  '_partner_qualification_',
  '_office_city_',
  '_office_state_',
  '_office_address_',
  '_office_zip_code_',
  '_office_total_value_',
  '_office_quotes_',
  '_office_quote_value_',
  '_partner_full_name_',
  '_parner_total_quotes_',      # Note: typo preserved
  '_partner_total_quotes_',     # Correct spelling
  '_partner_sum_',
  '_percentage_',
  '_total_quotes_',
  '_sum_percentage_'
]

puts "\n📋 Expected to replace #{expected_underline_placeholders.length} unique placeholder types"
puts "JSON data loaded: #{data['partners'].length} partners, #{data['society']['name']}"

# Open the underline template
original_path = 'tests/CS-TEMPLATE.docx'
output_path = 'tests/CS-TEMPLATE-validated-output.docx'

doc = Docx::Document.open(original_path)

puts "\n" + "-"*70
puts "STEP 1: PERFORMING INTELLIGENT REPLACEMENTS"
puts "-"*70

replacement_count = 0

# Process each paragraph with comprehensive error tracking
doc.paragraphs.each_with_index do |paragraph, para_index|
  original_text = paragraph.text
  
  # Track which placeholders we attempt to replace in this paragraph
  placeholders_in_paragraph = []
  
  # 1. OFFICE NAME
  paragraph.substitute_across_runs_with_block(/_office_name_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_name_'
    result = data['society']['name']
    puts "✅ Para #{para_index + 1}: _office_name_ → #{result[0..30]}..."
    result
  end
  
  # 2. PARTNER QUALIFICATION
  paragraph.substitute_across_runs_with_block(/_partner_qualification_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_partner_qualification_'
    partners = data['partners']
    
    if partners.length == 2
      p1, p2 = partners[0], partners[1]
      result = "#{p1['name']} (#{p1['oab_number']}) e #{p2['name']} (#{p2['oab_number']}), ambos Advogados"
    else
      result = "Multiple partners logic"
    end
    puts "✅ Para #{para_index + 1}: _partner_qualification_ → #{result[0..40]}..."
    result
  end
  
  # 3. OFFICE CITY
  paragraph.substitute_across_runs_with_block(/_office_city_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_city_'
    result = data['society']['city']
    puts "✅ Para #{para_index + 1}: _office_city_ → #{result}"
    result
  end
  
  # 4. OFFICE STATE
  paragraph.substitute_across_runs_with_block(/_office_state_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_state_'
    result = data['society']['state']
    puts "✅ Para #{para_index + 1}: _office_state_ → #{result}"
    result
  end
  
  # 5. OFFICE ADDRESS
  paragraph.substitute_across_runs_with_block(/_office_address_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_address_'
    result = data['society']['address']
    puts "✅ Para #{para_index + 1}: _office_address_ → #{result[0..30]}..."
    result
  end
  
  # 6. OFFICE ZIP CODE
  paragraph.substitute_across_runs_with_block(/_office_zip_code_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_zip_code_'
    result = data['society']['zip_code']
    puts "✅ Para #{para_index + 1}: _office_zip_code_ → #{result}"
    result
  end
  
  # 7. OFFICE TOTAL VALUE
  paragraph.substitute_across_runs_with_block(/_office_total_value_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_total_value_'
    total_capital = data['capital']['total_value'].to_f
    result = "#{total_capital.to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
    puts "✅ Para #{para_index + 1}: _office_total_value_ → R$ #{result}"
    result
  end
  
  # 8. OFFICE QUOTES
  paragraph.substitute_across_runs_with_block(/_office_quotes_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_quotes_'
    result = data['capital']['total_quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    puts "✅ Para #{para_index + 1}: _office_quotes_ → #{result}"
    result
  end
  
  # 9. OFFICE QUOTE VALUE
  paragraph.substitute_across_runs_with_block(/_office_quote_value_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_office_quote_value_'
    quote_value = data['capital']['quote_value'].to_f
    result = "#{quote_value.to_i},00"
    puts "✅ Para #{para_index + 1}: _office_quote_value_ → R$ #{result}"
    result
  end
  
  # 10. PARTNER FULL NAME
  paragraph.substitute_across_runs_with_block(/_partner_full_name_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_partner_full_name_'
    admin_partner = data['partners'].find { |p| p['is_administrator'] }
    result = admin_partner ? admin_partner['name'] : data['partners'].first['name']
    puts "✅ Para #{para_index + 1}: _partner_full_name_ → #{result}"
    result
  end
  
  # 11. PARTNER TOTAL QUOTES (with typo)
  paragraph.substitute_across_runs_with_block(/_parner_total_quotes_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_parner_total_quotes_'
    admin_partner = data['partners'].find { |p| p['is_administrator'] }
    partner_name = admin_partner ? admin_partner['name'] : data['partners'].first['name']
    partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
    result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    puts "✅ Para #{para_index + 1}: _parner_total_quotes_ → #{result}"
    result
  end
  
  # 12. PARTNER TOTAL QUOTES (correct)
  paragraph.substitute_across_runs_with_block(/_partner_total_quotes_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_partner_total_quotes_'
    admin_partner = data['partners'].find { |p| p['is_administrator'] }
    partner_name = admin_partner ? admin_partner['name'] : data['partners'].first['name']
    partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
    result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    puts "✅ Para #{para_index + 1}: _partner_total_quotes_ → #{result}"
    result
  end
  
  # 13. PARTNER SUM
  paragraph.substitute_across_runs_with_block(/_partner_sum_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_partner_sum_'
    admin_partner = data['partners'].find { |p| p['is_administrator'] }
    partner_name = admin_partner ? admin_partner['name'] : data['partners'].first['name']
    partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
    result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
    puts "✅ Para #{para_index + 1}: _partner_sum_ → R$ #{result}"
    result
  end
  
  # 14. PERCENTAGE
  paragraph.substitute_across_runs_with_block(/_percentage_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_percentage_'
    admin_partner = data['partners'].find { |p| p['is_administrator'] }
    partner_name = admin_partner ? admin_partner['name'] : data['partners'].first['name']
    partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
    result = "#{partner_capital['percentage']}%"
    puts "✅ Para #{para_index + 1}: _percentage_ → #{result}"
    result
  end
  
  # 15. TOTAL QUOTES
  paragraph.substitute_across_runs_with_block(/_total_quotes_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_total_quotes_'
    result = data['capital']['total_quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    puts "✅ Para #{para_index + 1}: _total_quotes_ → #{result}"
    result
  end
  
  # 16. SUM PERCENTAGE
  paragraph.substitute_across_runs_with_block(/_sum_percentage_/) do |match|
    replacement_count += 1
    placeholders_in_paragraph << '_sum_percentage_'
    total_percentage = data['capital']['partners'].sum { |p| p['percentage'] }
    result = "#{total_percentage}%"
    puts "✅ Para #{para_index + 1}: _sum_percentage_ → #{result}"
    result
  end
  
  # Log if this paragraph had content but no replacements
  if !original_text.strip.empty? && placeholders_in_paragraph.empty? && original_text.match?(/_\w+_/)
    puts "⚠️  Para #{para_index + 1}: Contains placeholders but no replacements made"
    puts "    Text: #{original_text[0..80]}..."
  end
end

puts "\n📊 Replacement Summary: #{replacement_count} total replacements made"

# Save the document
doc.save(output_path)

puts "\n" + "-"*70
puts "STEP 2: VALIDATING REPLACEMENTS"
puts "-"*70

# Validate the replacements
validator = Docx::ReplacementValidator.validate(original_path, output_path)
validator.report

# Additional specific validation
puts "\n" + "-"*70
puts "STEP 3: SPECIFIC PLACEHOLDER VALIDATION"
puts "-"*70

specific_validation = validator.validate_expected_placeholders(expected_underline_placeholders)

specific_validation.each do |result|
  status = result[:successfully_replaced] ? "✅" : "❌"
  puts "#{status} #{result[:expected]}: #{result[:found_in_original]} → #{result[:found_in_processed]}"
  
  if !result[:successfully_replaced] && result[:found_in_original] > 0
    puts "    ⚠️  Still found #{result[:found_in_processed]} instances in processed document"
    result[:locations_processed].each do |location|
      puts "       Para #{location[:paragraph_index]}: #{location[:context][0..50]}..."
    end
  end
end

puts "\n" + "="*70
puts "VALIDATION COMPLETE"
puts "="*70

if validator.passed?
  puts "\n🎉 SUCCESS: All placeholders were successfully replaced!"
  puts "📄 Validated output: #{output_path}"
else
  puts "\n❌ FAILURE: Some placeholders were not replaced"
  puts "🔧 Check the validation report above for details"
  puts "📄 Partial output: #{output_path}"
  
  # Exit with error code to indicate failure
  exit(1)
end

puts "\n✅ Document generation and validation completed successfully!"
puts "="*70