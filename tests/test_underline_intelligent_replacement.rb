#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'

puts "\n" + "="*70
puts "INTELLIGENT UNDERLINE REPLACEMENT - COMPLETE LEGAL DOCUMENT"
puts "="*70

# Load the JSON data
data = JSON.parse(File.read('tests/test_information.json'))

puts "\n📋 Loaded data:"
puts "  Society: #{data['society']['name']}"
puts "  Partners: #{data['partners'].length}"
puts "  Total Capital: R$ #{data['capital']['total_value'].to_f.to_i}"
puts "  Total Quotes: #{data['capital']['total_quotes']}"

# Open the underline template
doc = Docx::Document.open('tests/CS-TEMPLATE.docx')

puts "\n" + "-"*70
puts "PROCESSING INTELLIGENT UNDERLINE REPLACEMENTS"
puts "-"*70

# Create helper variables for calculations
total_capital = data['capital']['total_value'].to_f
total_quotes = data['capital']['total_quotes']
quote_value = data['capital']['quote_value'].to_f

# Process each paragraph with intelligent logic
doc.paragraphs.each do |paragraph|

  # 1. OFFICE NAME - Society name
  paragraph.substitute_across_runs_with_block(/_office_name_/) do |match|
    result = data['society']['name']
    puts "✅ Replaced _office_name_: #{result}"
    result
  end

  # 2. PARTNER QUALIFICATION - Complex logic for multiple partners
  paragraph.substitute_across_runs_with_block(/_partner_qualification_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner = partners.first
      result = "#{partner['profession']}, #{partner['oab_number']}"
    elsif partners.length == 2
      p1, p2 = partners[0], partners[1]
      result = "#{p1['name']} (#{p1['oab_number']}) e #{p2['name']} (#{p2['oab_number']}), ambos Advogados"
    else
      partner_list = partners.map { |p| "#{p['name']} (#{p['oab_number']})" }.join(', ')
      result = "#{partner_list}, todos Advogados"
    end
    puts "✅ Replaced _partner_qualification_: #{result[0..50]}..."
    result
  end

  # 3. OFFICE CITY
  paragraph.substitute_across_runs_with_block(/_office_city_/) do |match|
    result = data['society']['city']
    puts "✅ Replaced _office_city_: #{result}"
    result
  end

  # 4. OFFICE STATE
  paragraph.substitute_across_runs_with_block(/_office_state_/) do |match|
    result = data['society']['state']
    puts "✅ Replaced _office_state_: #{result}"
    result
  end

  # 5. OFFICE ADDRESS
  paragraph.substitute_across_runs_with_block(/_office_address_/) do |match|
    result = data['society']['address']
    puts "✅ Replaced _office_address_: #{result}"
    result
  end

  # 6. OFFICE ZIP CODE
  paragraph.substitute_across_runs_with_block(/_office_zip_code_/) do |match|
    result = data['society']['zip_code']
    puts "✅ Replaced _office_zip_code_: #{result}"
    result
  end

  # 7. OFFICE TOTAL VALUE - Format as Brazilian currency
  paragraph.substitute_across_runs_with_block(/_office_total_value_/) do |match|
    result = "#{total_capital.to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
    puts "✅ Replaced _office_total_value_: R$ #{result}"
    result
  end

  # 8. OFFICE QUOTES - Total number of quotes
  paragraph.substitute_across_runs_with_block(/_office_quotes_/) do |match|
    result = total_quotes.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    puts "✅ Replaced _office_quotes_: #{result}"
    result
  end

  # 9. OFFICE QUOTE VALUE - Individual quote value
  paragraph.substitute_across_runs_with_block(/_office_quote_value_/) do |match|
    result = "#{quote_value.to_i},00"
    puts "✅ Replaced _office_quote_value_: R$ #{result}"
    result
  end

  # 10. PARTNER FULL NAME - Intelligent logic for multiple partners
  paragraph.substitute_across_runs_with_block(/_partner_full_name_/) do |match|
    partners = data['partners']

    if partners.length == 1
      result = partners.first['name']
    else
      # For multiple partners, create a list or use administrator
      admin_partner = partners.find { |p| p['is_administrator'] }
      if admin_partner
        result = admin_partner['name']
      else
        # Use all partners
        result = partners.map { |p| p['name'] }.join(' e ')
      end
    end
    puts "✅ Replaced _partner_full_name_: #{result}"
    result
  end

  # 11. PARTNER TOTAL QUOTES (partner_total_quotes - with typo) - Smart logic based on context
  paragraph.substitute_across_runs_with_block(/_partner_total_quotes_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner_capital = data['capital']['partners'].first
      result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    else
      # If multiple partners, use administrator or first partner for this context
      admin_partner = partners.find { |p| p['is_administrator'] }
      partner_name = admin_partner ? admin_partner['name'] : partners.first['name']

      partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
      result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    end
    puts "✅ Replaced _partner_total_quotes_: #{result}"
    result
  end

  # 11b. PARTNER TOTAL QUOTES (correct spelling) - Same logic as above
  paragraph.substitute_across_runs_with_block(/_partner_total_quotes_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner_capital = data['capital']['partners'].first
      result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    else
      # If multiple partners, use administrator or first partner for this context
      admin_partner = partners.find { |p| p['is_administrator'] }
      partner_name = admin_partner ? admin_partner['name'] : partners.first['name']

      partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
      result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    end
    puts "✅ Replaced _partner_total_quotes_: #{result}"
    result
  end

  # 12. PARTNER SUM - Partner's total capital value
  paragraph.substitute_across_runs_with_block(/_partner_sum_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner_capital = data['capital']['partners'].first
      result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
    else
      # Use administrator or first partner
      admin_partner = partners.find { |p| p['is_administrator'] }
      partner_name = admin_partner ? admin_partner['name'] : partners.first['name']

      partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
      result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
    end
    puts "✅ Replaced _partner_sum_: R$ #{result}"
    result
  end

  # 13. TOTAL QUOTES - Same as office_quotes
  paragraph.substitute_across_runs_with_block(/_total_quotes_/) do |match|
    result = total_quotes.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    puts "✅ Replaced _total_quotes_: #{result}"
    result
  end

  # 14. PERCENTAGE - Partner's ownership percentage
  paragraph.substitute_across_runs_with_block(/_percentage_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner_capital = data['capital']['partners'].first
      result = "#{partner_capital['percentage']}%"
    else
      # Use administrator or first partner
      admin_partner = partners.find { |p| p['is_administrator'] }
      partner_name = admin_partner ? admin_partner['name'] : partners.first['name']

      partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
      result = "#{partner_capital['percentage']}%"
    end
    puts "✅ Replaced _percentage_: #{result}"
    result
  end

  # 15. SUM PERCENTAGE - Total should always be 100%
  paragraph.substitute_across_runs_with_block(/_sum_percentage_/) do |match|
    total_percentage = data['capital']['partners'].sum { |p| p['percentage'] }
    result = "#{total_percentage}%"
    puts "✅ Replaced _sum_percentage_: #{result}"
    result
  end

end

puts "\n" + "-"*70
puts "PROCESSING TABLES"
puts "-"*70

# Process all tables in the document
doc.tables.each_with_index do |table, table_index|
  puts "\n📊 Processing Table #{table_index + 1}"
  
  table.rows.each_with_index do |row, row_index|
    row.cells.each_with_index do |cell, cell_index|
      cell.paragraphs.each_with_index do |paragraph, para_index|
        
        # Apply the same replacement logic to table cells
        
        # 1. OFFICE NAME
        paragraph.substitute_across_runs_with_block(/_office_name_/) do |match|
          result = data['society']['name']
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_name_ → #{result[0..20]}..."
          result
        end
        
        # 2. PARTNER QUALIFICATION
        paragraph.substitute_across_runs_with_block(/_partner_qualification_/) do |match|
          partners = data['partners']
          
          if partners.length == 1
            partner = partners.first
            result = "#{partner['profession']}, #{partner['oab_number']}"
          elsif partners.length == 2
            p1, p2 = partners[0], partners[1]
            result = "#{p1['name']} (#{p1['oab_number']}) e #{p2['name']} (#{p2['oab_number']}), ambos Advogados"
          else
            partner_list = partners.map { |p| "#{p['name']} (#{p['oab_number']})" }.join(', ')
            result = "#{partner_list}, todos Advogados"
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_qualification_ → #{result[0..30]}..."
          result
        end
        
        # 3. PARTNER FULL NAME
        paragraph.substitute_across_runs_with_block(/_partner_full_name_/) do |match|
          partners = data['partners']
          
          if partners.length == 1
            result = partners.first['name']
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            if admin_partner
              result = admin_partner['name']
            else
              result = partners.map { |p| p['name'] }.join(' e ')
            end
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_full_name_ → #{result}"
          result
        end
        
        # 4. PARTNER TOTAL QUOTES (with typo)
        paragraph.substitute_across_runs_with_block(/_parner_total_quotes_/) do |match|
          partners = data['partners']
          
          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_name = admin_partner ? admin_partner['name'] : partners.first['name']
            
            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
            result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _parner_total_quotes_ → #{result}"
          result
        end
        
        # 5. PARTNER TOTAL QUOTES (correct spelling)
        paragraph.substitute_across_runs_with_block(/_partner_total_quotes_/) do |match|
          partners = data['partners']
          
          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_name = admin_partner ? admin_partner['name'] : partners.first['name']
            
            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
            result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_total_quotes_ → #{result}"
          result
        end
        
        # 6. PARTNER SUM
        paragraph.substitute_across_runs_with_block(/_partner_sum_/) do |match|
          partners = data['partners']
          
          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_name = admin_partner ? admin_partner['name'] : partners.first['name']
            
            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
            result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_sum_ → R$ #{result}"
          result
        end
        
        # 7. PERCENTAGE
        paragraph.substitute_across_runs_with_block(/_percentage_/) do |match|
          partners = data['partners']
          
          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = "#{partner_capital['percentage']}%"
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_name = admin_partner ? admin_partner['name'] : partners.first['name']
            
            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_name }
            result = "#{partner_capital['percentage']}%"
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _percentage_ → #{result}"
          result
        end
        
        # 8. TOTAL QUOTES
        paragraph.substitute_across_runs_with_block(/_total_quotes_/) do |match|
          result = total_quotes.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _total_quotes_ → #{result}"
          result
        end
        
        # 9. SUM PERCENTAGE
        paragraph.substitute_across_runs_with_block(/_sum_percentage_/) do |match|
          total_percentage = data['capital']['partners'].sum { |p| p['percentage'] }
          result = "#{total_percentage}%"
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _sum_percentage_ → #{result}"
          result
        end
        
        # 10. OFFICE CITY
        paragraph.substitute_across_runs_with_block(/_office_city_/) do |match|
          result = data['society']['city']
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_city_ → #{result}"
          result
        end
        
        # 11. OFFICE STATE
        paragraph.substitute_across_runs_with_block(/_office_state_/) do |match|
          result = data['society']['state']
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_state_ → #{result}"
          result
        end
        
        # 12. OFFICE ADDRESS
        paragraph.substitute_across_runs_with_block(/_office_address_/) do |match|
          result = data['society']['address']
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_address_ → #{result[0..30]}..."
          result
        end
        
        # 13. OFFICE ZIP CODE
        paragraph.substitute_across_runs_with_block(/_office_zip_code_/) do |match|
          result = data['society']['zip_code']
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_zip_code_ → #{result}"
          result
        end
        
        # 14. OFFICE TOTAL VALUE
        paragraph.substitute_across_runs_with_block(/_office_total_value_/) do |match|
          result = "#{total_capital.to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_total_value_ → R$ #{result}"
          result
        end
        
        # 15. OFFICE QUOTES
        paragraph.substitute_across_runs_with_block(/_office_quotes_/) do |match|
          result = total_quotes.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_quotes_ → #{result}"
          result
        end
        
        # 16. OFFICE QUOTE VALUE
        paragraph.substitute_across_runs_with_block(/_office_quote_value_/) do |match|
          result = "#{quote_value.to_i},00"
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _office_quote_value_ → R$ #{result}"
          result
        end
        
      end
    end
  end
end

# Save the intelligent document
output_path = 'tests/CS-TEMPLATE-underline-intelligent-output.docx'
doc.save(output_path)

puts "\n" + "="*70
puts "INTELLIGENT UNDERLINE REPLACEMENT COMPLETE"
puts "="*70

puts "\n📄 Generated document: #{output_path}"
puts "\n🎯 Replacement Summary:"
puts "  Society Information:"
puts "    - Name: #{data['society']['name']}"
puts "    - Location: #{data['society']['city']}, #{data['society']['state']}"
puts "    - Address: #{data['society']['address']}"

puts "\n  Capital Structure:"
puts "    - Total Value: R$ #{total_capital.to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
puts "    - Total Quotes: #{total_quotes.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')}"
puts "    - Quote Value: R$ #{quote_value.to_i},00"

puts "\n  Partners:"
data['capital']['partners'].each do |partner|
  puts "    - #{partner['name']}: #{partner['quotes']} quotes (#{partner['percentage']}%)"
end

puts "\n💡 Intelligent Logic Applied:"
puts "  - Multi-partner qualification formatting"
puts "  - Administrator preference for individual partner fields"
puts "  - Brazilian currency formatting (R$ 10.000,00)"
puts "  - Automatic percentage calculations"
puts "  - Context-aware partner selection"

puts "\n" + "="*70
