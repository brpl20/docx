#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'

puts "\n" + "="*70
puts "INTELLIGENT JSON-BASED DOCUMENT GENERATION"
puts "="*70

# Load the JSON data
data = JSON.parse(File.read('tests/test_information.json'))

puts "\n📋 Loaded data:"
puts "  Society: #{data['society']['name']}"
puts "  Partners: #{data['partners'].length}"
data['partners'].each_with_index do |partner, i|
  puts "    #{i+1}. #{partner['name']} (#{partner['oab_number']})"
end

# Open the mustache template
doc = Docx::Document.open('tests/CS-TEMPLATE-mus.docx')

puts "\n" + "-"*70
puts "PROCESSING INTELLIGENT REPLACEMENTS"
puts "-"*70

# Process each paragraph with intelligent logic
doc.paragraphs.each do |paragraph|
  
  # 1. OFFICE NAME - Simple replacement
  paragraph.substitute_across_runs_with_block(/\{\s*office_name\s*\}/) do |match|
    result = data['society']['name']
    puts "✅ Replaced office_name: #{result}"
    result
  end
  
  # 2. PARTNER QUALIFICATION - Complex logic for multiple partners
  paragraph.substitute_across_runs_with_block(/\{\s*partner_qualification\s*\}/) do |match|
    partners = data['partners']
    
    if partners.length == 1
      # Single partner
      partner = partners.first
      result = "#{partner['profession']}, #{partner['oab_number']}"
      puts "✅ Single partner qualification: #{result}"
      result
    elsif partners.length == 2
      # Two partners - merge their qualifications
      p1, p2 = partners[0], partners[1]
      result = "#{p1['name']} (#{p1['oab_number']}) e #{p2['name']} (#{p2['oab_number']}), ambos Advogados"
      puts "✅ Dual partner qualification: #{result}"
      result
    else
      # Multiple partners - create list
      partner_list = partners.map { |p| "#{p['name']} (#{p['oab_number']})" }.join(', ')
      result = "#{partner_list}, todos Advogados"
      puts "✅ Multiple partner qualification: #{result}"
      result
    end
  end
  
  # 3. OFFICE ADDRESS - Combine address components
  paragraph.substitute_across_runs_with_block(/\{\s*office_address\s*\}/) do |match|
    society = data['society']
    result = society['address']
    puts "✅ Replaced office_address: #{result}"
    result
  end
  
  # 4. OFFICE CITY
  paragraph.substitute_across_runs_with_block(/\{\s*office_city\s*\}/) do |match|
    result = data['society']['city']
    puts "✅ Replaced office_city: #{result}"
    result
  end
  
  # 5. OFFICE STATE
  paragraph.substitute_across_runs_with_block(/\{\s*office_state\s*\}/) do |match|
    result = data['society']['state']
    puts "✅ Replaced office_state: #{result}"
    result
  end
  
  # 6. OFFICE ZIP CODE
  paragraph.substitute_across_runs_with_block(/\{\s*office_zip_code\s*\}/) do |match|
    result = data['society']['zip_code']
    puts "✅ Replaced office_zip_code: #{result}"
    result
  end
  
end

# Save the intelligent document
output_path = 'tests/CS-TEMPLATE-intelligent-output.docx'
doc.save(output_path)

puts "\n" + "="*70
puts "INTELLIGENT REPLACEMENT COMPLETE"
puts "="*70

puts "\n📄 Generated document: #{output_path}"
puts "\n🎯 Key Features Demonstrated:"
puts "  1. ✅ JSON data loading and parsing"
puts "  2. ✅ Simple field replacement (office_name)"
puts "  3. ✅ Complex conditional logic (partner_qualification)"
puts "  4. ✅ Multiple partner handling with proper formatting"
puts "  5. ✅ Address component mapping"
puts "  6. ✅ Fragmentation-resistant replacement across all fields"

puts "\n💡 Partner Qualification Logic:"
if data['partners'].length == 1
  puts "  → Single partner: Shows profession and OAB number"
elsif data['partners'].length == 2
  puts "  → Two partners: 'Partner1 (OAB1) e Partner2 (OAB2), ambos Advogados'"
else
  puts "  → Multiple partners: Lists all with 'todos Advogados'"
end

puts "\n🔄 Next Steps:"
puts "  - Add logic for capital information"
puts "  - Handle partner-specific data (quotes, percentages)"
puts "  - Add administrator designation logic"
puts "  - Implement table generation for multiple partners"

puts "\n" + "="*70