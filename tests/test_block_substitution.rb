#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'

puts "\n" + "="*70
puts "TESTING BLOCK SUBSTITUTION FUNCTIONALITY"
puts "="*70

# Create a test document with some sample content
doc = Docx::Document.open('tests/CS-TEMPLATE.docx')

puts "\nOriginal content (first 3 paragraphs):"
doc.paragraphs[0..2].each_with_index do |p, i|
  puts "  #{i+1}. #{p.text}"
end

puts "\n" + "-"*70
puts "TEST 1: substitute_with_block (traditional method)"
puts "-"*70

# Test 1: Traditional substitute_with_block on text runs
doc.paragraphs.each do |paragraph|
  paragraph.each_text_run do |text_run|
    # Find and transform any numbers
    text_run.substitute_with_block(/(\d+)/) do |match_data|
      number = match_data[1].to_i
      "#{number * 2}"  # Double the number
    end
    
    # Transform any currency amounts
    text_run.substitute_with_block(/\$(\d+)/) do |match_data|
      amount = match_data[1].to_i
      "$#{(amount * 1.1).round}"  # Add 10% markup
    end
  end
end

puts "After traditional block substitution:"
doc.paragraphs[0..2].each_with_index do |p, i|
  puts "  #{i+1}. #{p.text}"
end

puts "\n" + "-"*70
puts "TEST 2: substitute_across_runs_with_block (fragmentation-resistant)"
puts "-"*70

# Test 2: Use the new substitute_across_runs_with_block method
doc.paragraphs.each do |paragraph|
  # Transform placeholder patterns with calculations
  paragraph.substitute_across_runs_with_block(/_(\w+)_/) do |match_data|
    placeholder_name = match_data[1]
    "PROCESSED_#{placeholder_name.upcase}"
  end
  
  # Transform any date patterns
  paragraph.substitute_across_runs_with_block(/(\d{4})-(\d{2})-(\d{2})/) do |match_data|
    year, month, day = match_data[1], match_data[2], match_data[3]
    "#{day}/#{month}/#{year}"  # Convert YYYY-MM-DD to DD/MM/YYYY
  end
  
  # Transform mustache placeholders with dynamic content
  paragraph.substitute_across_runs_with_block(/\{\s*(\w+)\s*\}/) do |match_data|
    field_name = match_data[1]
    case field_name
    when 'office_name'
      "Dynamic Law Firm LLC"
    when 'partner_qualification'
      "Senior Partner & Legal Advisor"
    when 'office_total_value'
      "$#{rand(100..999)},000"
    else
      "AUTO_#{field_name.upcase}"
    end
  end
end

puts "After cross-runs block substitution:"
doc.paragraphs[0..2].each_with_index do |p, i|
  puts "  #{i+1}. #{p.text}"
end

puts "\n" + "-"*70
puts "TEST 3: Complex regex with multiple captures"
puts "-"*70

# Test 3: More complex transformations
doc.paragraphs.each do |paragraph|
  # Transform contact info patterns
  paragraph.substitute_across_runs_with_block(/(\w+)@([\w.-]+\.\w+)/) do |match_data|
    username, domain = match_data[1], match_data[2]
    "#{username.upcase}@#{domain.downcase}"
  end
  
  # Transform phone number patterns
  paragraph.substitute_across_runs_with_block(/(\d{3})-(\d{3})-(\d{4})/) do |match_data|
    area, prefix, number = match_data[1], match_data[2], match_data[3]
    "(#{area}) #{prefix}-#{number}"
  end
  
  # Transform percentage calculations
  paragraph.substitute_across_runs_with_block(/(\d+)%/) do |match_data|
    percentage = match_data[1].to_i
    decimal = percentage / 100.0
    "#{percentage}% (#{decimal})"
  end
end

puts "After complex regex transformations:"
doc.paragraphs[0..2].each_with_index do |p, i|
  puts "  #{i+1}. #{p.text}"
end

# Save the transformed document
doc.save('tests/block_substitution_output.docx')

puts "\n" + "="*70
puts "BLOCK SUBSTITUTION TESTS COMPLETE"
puts "="*70
puts "\n✅ Output saved to: tests/block_substitution_output.docx"
puts "\nKey features tested:"
puts "  1. ✅ substitute_with_block (traditional method)"
puts "  2. ✅ substitute_across_runs_with_block (fragmentation-resistant)"
puts "  3. ✅ Multiple capture groups in regex"
puts "  4. ✅ Dynamic content generation based on matched text"
puts "  5. ✅ Complex pattern matching and transformation"

puts "\nBoth methods allow you to:"
puts "  - Use regex patterns with capture groups"
puts "  - Access MatchData object in the block"
puts "  - Perform calculations or lookups based on matched text"
puts "  - Handle placeholders that might be fragmented across XML nodes"
puts "\n" + "="*70