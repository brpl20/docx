#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'
require 'date'

puts "\n" + "="*70
puts "INDIVIDUAL LAWYER TEMPLATE - DOCUMENT PROCESSING"
puts "="*70

# Load the JSON data for individual lawyer
data = JSON.parse(File.read('tests/test_individual_information.json'))

puts "\n📋 Loaded data:"
puts "  Office: #{data['society']['name']}"
puts "  Lawyer: #{data['partner']['name']} #{data['partner']['last_name']}"
puts "  Total Capital: R$ #{data['capital']['total_value'].to_f.to_i}"
puts "  Total Quotes: #{data['capital']['total_quotes']}"

# Open the individual template
doc = Docx::Document.open('tests/CS-UNIPESSOAL-TEMPLATE.docx')

# Helper methods
def full_name(p)
  [p['name'], p['last_name']].compact.join(' ').squeeze(' ').strip
end

def address_str(p)
  base = [p['address'], p['number']].compact.join(', ').strip
  base += " - #{p['complement']}" if p['complement'] && !p['complement'].empty?
  tail = [
    p['neighborhood'],
    [p['city'], p['state']].compact.join(' - '),
    p['zip_code'] ? "CEP #{p['zip_code']}" : nil
  ].compact.join(', ')
  [base, tail].reject(&:empty?).join(', ')
end

def qualification(p)
  "#{full_name(p)}, #{p['nationality']}, #{p['civil_status']}, #{p['profession']}, " \
  "inscrito(a) na #{p['oab_number']}, CPF #{p['cpf']}, nascido(a) em #{p['birth_city']} " \
  "em #{p['birth_date']}, residente e domiciliado(a) à #{address_str(p)}"
end

# Format currency in Brazilian format
def format_currency(value)
  value.to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
end

puts "\n" + "-"*70
puts "PROCESSING DOCUMENT - REPLACING PLACEHOLDERS"
puts "-"*70

# Process each paragraph
doc.paragraphs.each do |paragraph|
  
  # 1. PARTNER QUALIFICATION
  paragraph.substitute_across_runs_with_block(/_partner_qualification_/) do |match|
    result = qualification(data['partner'])
    puts "✅ Replaced _partner_qualification_: #{result[0..80]}..."
    result
  end
  
  # 2. PARTNER FULL NAME
  paragraph.substitute_across_runs_with_block(/_partner_full_name_/) do |match|
    result = full_name(data['partner'])
    puts "✅ Replaced _partner_full_name_: #{result}"
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
  
  # 7. OFFICE TOTAL VALUE
  paragraph.substitute_across_runs_with_block(/_office_total_value_/) do |match|
    result = "R$ #{format_currency(data['capital']['total_value'])},00"
    puts "✅ Replaced _office_total_value_: #{result}"
    result
  end
  
  # 8. OFFICE QUOTES
  paragraph.substitute_across_runs_with_block(/_office_quotes_/) do |match|
    result = format_currency(data['capital']['total_quotes'])
    puts "✅ Replaced _office_quotes_: #{result}"
    result
  end
  
  # 9. OFFICE QUOTE VALUE  
  paragraph.substitute_across_runs_with_block(/_office_quote_value_/) do |match|
    result = "R$ #{data['capital']['quote_value'].to_i},00"
    puts "✅ Replaced _office_quote_value_: #{result}"
    result
  end
  
  # 10. DATE - Brazilian format
  paragraph.substitute_across_runs_with_block(/_date_/) do |match|
    today = Date.today
    months = [
      'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
      'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
    ]
    result = "#{today.day} de #{months[today.month - 1]} de #{today.year}"
    puts "✅ Replaced _date_: #{result}"
    result
  end
  
  # 11. PARTNER FULL NAME (with typo - parner instead of partner)
  paragraph.substitute_across_runs_with_block(/_parner_full_name_/) do |match|
    result = full_name(data['partner'])
    puts "✅ Replaced _parner_full_name_: #{result}"
    result
  end
end

# Process tables if any
doc.tables.each_with_index do |table, table_index|
  puts "\n📊 Processing Table #{table_index + 1}"
  
  table.rows.each_with_index do |row, row_index|
    row.cells.each_with_index do |cell, cell_index|
      cell.paragraphs.each do |paragraph|
        
        # Same replacements for table cells
        paragraph.substitute_across_runs_with_block(/_partner_qualification_/) do |match|
          result = qualification(data['partner'])
          puts "✅ Table #{table_index + 1}: Replaced _partner_qualification_"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_partner_full_name_/) do |match|
          result = full_name(data['partner'])
          puts "✅ Table #{table_index + 1}: Replaced _partner_full_name_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_city_/) do |match|
          result = data['society']['city']
          puts "✅ Table #{table_index + 1}: Replaced _office_city_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_state_/) do |match|
          result = data['society']['state']
          puts "✅ Table #{table_index + 1}: Replaced _office_state_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_address_/) do |match|
          result = data['society']['address']
          puts "✅ Table #{table_index + 1}: Replaced _office_address_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_zip_code_/) do |match|
          result = data['society']['zip_code']
          puts "✅ Table #{table_index + 1}: Replaced _office_zip_code_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_total_value_/) do |match|
          result = "R$ #{format_currency(data['capital']['total_value'])},00"
          puts "✅ Table #{table_index + 1}: Replaced _office_total_value_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_quotes_/) do |match|
          result = format_currency(data['capital']['total_quotes'])
          puts "✅ Table #{table_index + 1}: Replaced _office_quotes_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_office_quote_value_/) do |match|
          result = "R$ #{data['capital']['quote_value'].to_i},00"
          puts "✅ Table #{table_index + 1}: Replaced _office_quote_value_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_date_/) do |match|
          today = Date.today
          months = [
            'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
            'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
          ]
          result = "#{today.day} de #{months[today.month - 1]} de #{today.year}"
          puts "✅ Table #{table_index + 1}: Replaced _date_: #{result}"
          result
        end
        
        paragraph.substitute_across_runs_with_block(/_parner_full_name_/) do |match|
          result = full_name(data['partner'])
          puts "✅ Table #{table_index + 1}: Replaced _parner_full_name_: #{result}"
          result
        end
      end
    end
  end
end

# Save the result
output_file = 'tests/CS-UNIPESSOAL-output.docx'
doc.save(output_file)

puts "\n" + "="*70
puts "INDIVIDUAL TEMPLATE PROCESSING COMPLETE"
puts "="*70
puts "\n📄 Generated document: #{output_file}"
puts "\n🎯 Replacement Summary:"
puts "  Office Information:"
puts "    - Name: #{data['society']['name']}"
puts "    - Location: #{data['society']['city']}, #{data['society']['state']}"
puts "    - Address: #{data['society']['address']}"
puts "    - ZIP: #{data['society']['zip_code']}"
puts ""
puts "  Lawyer Information:"
puts "    - Name: #{full_name(data['partner'])}"
puts "    - CPF: #{data['partner']['cpf']}"
puts "    - OAB: #{data['partner']['oab_number']}"
puts ""
puts "  Capital Structure:"
puts "    - Total Value: R$ #{format_currency(data['capital']['total_value'])},00"
puts "    - Total Quotes: #{format_currency(data['capital']['total_quotes'])}"
puts "    - Quote Value: R$ #{data['capital']['quote_value'].to_i},00"
puts "\n" + "="*70