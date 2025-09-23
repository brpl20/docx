#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'

puts "\n" + "="*70
puts "COMPARING BLOCK SUBSTITUTION METHODS"
puts "="*70

# Test with both template files to see the difference
templates = ['tests/CS-TEMPLATE.docx', 'tests/CS-TEMPLATE-mus.docx']

templates.each do |template_path|
  next unless File.exist?(template_path)
  
  puts "\n" + "-"*70
  puts "Testing: #{File.basename(template_path)}"
  puts "-"*70
  
  doc = Docx::Document.open(template_path)
  
  # Show original content with placeholders
  puts "\nOriginal placeholders found:"
  doc.paragraphs.each_with_index do |p, i|
    text = p.text
    if text.match?(/_\w+_/) || text.match?(/\{\s*\w+\s*\}/)
      puts "  Paragraph #{i+1}: #{text}"
    end
  end
  
  # Method 1: Traditional substitute_with_block
  puts "\n1. Traditional substitute_with_block:"
  success_count = 0
  doc.paragraphs.each do |paragraph|
    paragraph.each_text_run do |text_run|
      original = text_run.text
      
      # Try to replace underline placeholders
      text_run.substitute_with_block(/_(\w+)_/) do |match_data|
        success_count += 1
        "TRADITIONAL_#{match_data[1].upcase}"
      end
      
      # Try to replace mustache placeholders
      text_run.substitute_with_block(/\{\s*(\w+)\s*\}/) do |match_data|
        success_count += 1
        "TRADITIONAL_#{match_data[1].upcase}"
      end
      
      if text_run.text != original
        puts "    ✅ Replaced in text run: #{text_run.text}"
      end
    end
  end
  puts "    Total replacements: #{success_count}"
  
  # Reset document for second test
  doc = Docx::Document.open(template_path)
  
  # Method 2: Fragmentation-resistant substitute_across_runs_with_block
  puts "\n2. Fragmentation-resistant substitute_across_runs_with_block:"
  success_count = 0
  doc.paragraphs.each do |paragraph|
    original = paragraph.text
    
    # Replace underline placeholders
    paragraph.substitute_across_runs_with_block(/_(\w+)_/) do |match_data|
      success_count += 1
      "FRAGPROOF_#{match_data[1].upcase}"
    end
    
    # Replace mustache placeholders  
    paragraph.substitute_across_runs_with_block(/\{\s*(\w+)\s*\}/) do |match_data|
      success_count += 1
      "FRAGPROOF_#{match_data[1].upcase}"
    end
    
    if paragraph.text != original
      puts "    ✅ Replaced in paragraph: #{paragraph.text}"
    end
  end
  puts "    Total replacements: #{success_count}"
  
  # Save the result
  output_name = "tests/#{File.basename(template_path, '.docx')}_block_comparison.docx"
  doc.save(output_name)
  puts "\n📄 Saved result: #{output_name}"
end

puts "\n" + "="*70
puts "COMPARISON COMPLETE"
puts "="*70

puts "\n📋 Summary:"
puts "1. substitute_with_block (traditional):"
puts "   - Works on individual text runs"
puts "   - May miss placeholders fragmented across runs"
puts "   - Good for simple documents"

puts "\n2. substitute_across_runs_with_block (new):"
puts "   - Works across entire paragraph text"
puts "   - Handles fragmented placeholders"
puts "   - More reliable for complex documents"

puts "\n💡 Use case examples:"
puts <<~EXAMPLES

# Example 1: Simple number doubling
paragraph.substitute_across_runs_with_block(/total: (\\d+)/) do |match|
  "total: \#{match[1].to_i * 2}"
end

# Example 2: Dynamic placeholder replacement
paragraph.substitute_across_runs_with_block(/\\{\\s*(\\w+)\\s*\\}/) do |match|
  case match[1]
  when 'date'
    Date.today.to_s
  when 'amount'
    "$\#{rand(1000..9999)}"
  else
    "UNKNOWN_\#{match[1].upcase}"
  end
end

# Example 3: Complex calculations
paragraph.substitute_across_runs_with_block(/(\\d+)%/) do |match|
  percentage = match[1].to_i
  "\#{percentage}% (\#{(percentage/100.0).round(2)} decimal)"
end

EXAMPLES

puts "="*70