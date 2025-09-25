#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'
require 'date'

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

# Select Template
partners = data['partners']
template_file = partners.length == 1 ? 'tests/CS-TEMPLATE-INDIVIDUAL.docx' : 'tests/CS-TEMPLATE.docx'
doc = Docx::Document.open(template_file)

# Helpers
## Full Name
def full_name(p)
  [p['name'], p['last_name']].compact.join(' ').squeeze(' ').strip
end


## Address
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

## Qualification
def qualification(p)
  "#{full_name(p)}, #{p['nationality']}, #{p['civil_status']}, #{p['profession']}, " \
  "inscrito(a) na #{p['oab_number']}, CPF #{p['cpf']}, nascido(a) em #{p['birth_city']} " \
  "em #{p['birth_date']}, residente e domiciliado(a) à #{address_str(p)}"
end

## Calculations
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
    quals = partners.map { |p| qualification(p) }

    result =
      if partners.length == 1
        quals.first
      elsif partners.length == 2
        quals.join(' e ')
      else
        quals.join('; ')
      end

    puts "✅ Replaced _partner_qualification_: #{result[0..80]}..."
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

  # 11. PARTNER TOTAL QUOTES (partner_subscription) - Smart logic based on context
  paragraph.substitute_across_runs_with_block(/_partner_subscription_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner_capital = data['capital']['partners'].first
      result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    else
      # Multiple Partners Logic ->
      # For each partner we must insert a new line
      # We must make a composition of elements to get this
      # Element 1: Name
      # Element 2: Total Quotes -> _partner_total_quotes_
      # Element 3: Quotes Value -> _office_quote_value_
      # Element 4: Total -> _partner_sum_
      # -> Create a new line here and start over

      # Build multi-line result for all partners
      partner_lines = []
      data['capital']['partners'].each_with_index do |partner_capital, index|
        # Match by comparing full names
        partner_info = partners.find { |p| full_name(p) == partner_capital['name'] }
        next unless partner_info

        partner_name = partner_capital['name']  # Use the name from capital data
        partner_quotes = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
        partner_quote_value = "#{quote_value.to_i},00"
        partner_total = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"

        line = "O Sócio #{partner_name}, subscreve e integraliza neste ato #{partner_quotes} quotas no valor de R$ #{partner_quote_value} cada uma, perfazendo o total de R$ #{partner_total};"
        partner_lines << line
      end

      # Try using double newlines for better paragraph separation
      result = partner_lines.join("\n\n")
      puts "-----------------------"
      puts result
    end
    puts "✅ Replaced _partner_subscription_: #{result}"
    result
  end

  # 11b. PARTNER TOTAL QUOTES (correct spelling) - Same logic as above
  paragraph.substitute_across_runs_with_block(/_partner_total_quotes_/) do |match|
    partners = data['partners']

    if partners.length == 1
      partner_capital = data['capital']['partners'].first
      result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
    else
      # For multiple partners, create a multi-line result with each partner's quotes
      partner_lines = []
      data['capital']['partners'].each do |partner_capital|
        # Match by comparing full names
        partner_info = partners.find { |p| full_name(p) == partner_capital['name'] }
        next unless partner_info

        partner_name = partner_capital['name']  # Use the name from capital data
        partner_quotes = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')

        line = "#{partner_name}: #{partner_quotes} quotas"
        partner_lines << line
      end

      result = partner_lines.join("\n")
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
      # For multiple partners, create a multi-line result with each partner's total
      partner_lines = []
      data['capital']['partners'].each do |partner_capital|
        # Match by comparing full names
        partner_info = partners.find { |p| full_name(p) == partner_capital['name'] }
        next unless partner_info

        partner_name = partner_capital['name']  # Use the name from capital data
        partner_total = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"

        line = "#{partner_name}: R$ #{partner_total}"
        partner_lines << line
      end

      result = partner_lines.join("\n")
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

  # 16. PRO LABORE - Replace with "Parágrafo Sétimo:" or remove if false/nil
  paragraph.substitute_across_runs_with_block(/(?<![_\w])_pro_labore_(?![_\w])/) do |match|
    if data['society']['pro_labore']
      result = "Parágrafo Sétimo:"
      puts "✅ Replaced _pro_labore_: #{result}"
      result
    else
      puts "✅ Removed _pro_labore_ (pro_labore is false/nil)"
      ""
    end
  end

  # 17. PRO LABORE TEXT - Replace with specific text or remove if false/nil
  paragraph.substitute_across_runs_with_block(/(?<![_\w])_pro_labore_text_(?![_\w])/) do |match|
    if data['society']['pro_labore']
      result = "Pelo exercício da administração terão os sócios administradores direito a uma retirada mensal a título de \"pró-labore\", cujo valor será fixado em comum acordo entre os sócios e levado à conta de Despesas Gerais da Sociedade."
      puts "✅ Replaced _pro_labore_text_: #{result[0..80]}..."
      result
    else
      puts "✅ Removed _pro_labore_text_ (pro_labore is false/nil)"
      ""
    end
  end

  # 18. DIVIDENDS - Replace with "Parágrafo Terceiro:" or remove if false/nil
  paragraph.substitute_across_runs_with_block(/(?<![_\w])_dividends_(?![_\w])/) do |match|
    if data['society']['dividends']
      result = "Parágrafo Terceiro:"
      puts "✅ Replaced _dividends_: #{result}"
      result
    else
      puts "✅ Removed _dividends_ (dividends is false/nil)"
      ""
    end
  end

  # 19. DIVIDENDS TEXT - Replace with pattern-based text or remove if false/nil
  paragraph.substitute_across_runs_with_block(/(?<![_\w])_dividends_text_(?![_\w])/) do |match|
    if data['society']['dividends']
      pattern = data['society']['dividends_pattern'] || 1
      result = case pattern
               when 1
                 "Os sócios receberão dividendos proporcionais às suas respectivas participações no capital social."
               when 2
                 "Os sócios receberão dividendos desproporcionais às suas respectivas participações no capital social."
               else
                 "Os sócios receberão dividendos proporcionais às suas respectivas participações no capital social."
               end
      puts "✅ Replaced _dividends_text_ (pattern #{pattern}): #{result[0..80]}..."
      result
    else
      puts "✅ Removed _dividends_text_ (dividends is false/nil)"
      ""
    end
  end

  # 20. DATE - Replace with today's date in Brazilian format
  paragraph.substitute_across_runs_with_block(/(?<![_\w])_date_(?![_\w])/) do |match|
    today = Date.today
    months = [
      'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
      'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
    ]
    result = "#{today.day} de #{months[today.month - 1]} de #{today.year}"
    puts "✅ Replaced _date_: #{result}"
    result
  end

end

puts "\n" + "-"*70
puts "PROCESSING TABLES - STEP 1: ADD ROWS"
puts "-"*70

# First pass: Add rows for multiple partners
doc.tables.each_with_index do |table, table_index|
  puts "\n📊 Table #{table_index + 1}: Adding rows for multiple partners"

  # Find the partner row
  partner_row_index = nil

  table.rows.each_with_index do |row, row_index|
    row_text = row.cells.map { |cell|
      cell.paragraphs.map(&:to_s).join(' ')
    }.join(' ')

    if row_text.include?('_partner_full_name_') ||
       row_text.include?('_partner_total_quotes_') ||
       row_text.include?('_partner_sum_')
      partner_row_index = row_index
      puts "  Found partner row at index: #{partner_row_index}"
      break
    end
  end

  # Add rows if we have multiple partners (Table 1 - Capital table)
  if partner_row_index && data['partners'].length > 1
    partner_row = table.rows[partner_row_index]
    partner_row_node = partner_row.node

    num_rows_to_add = data['partners'].length - 1
    puts "  Adding #{num_rows_to_add} new row(s) for capital table..."

    last_inserted = partner_row_node
    num_rows_to_add.times do |i|
      partner_number = i + 2  # Start from 2 (first partner uses original placeholders)
      new_row = partner_row_node.dup
      # ??? DUP ?

      # The placeholders are split across multiple text runs due to formatting
      # We need to handle this more carefully to preserve the structure

      # For each cell in the new row, we need to modify placeholders while preserving formatting
      cells = new_row.xpath('.//w:tc')
      cells.each_with_index do |cell, cell_idx|
        # Create a temporary paragraph object to use the substitute_across_runs_with_block method
        cell.xpath('.//w:p').each do |p_node|
          temp_paragraph = Docx::Elements::Containers::Paragraph.new(p_node, {}, nil)

          # Replace each placeholder with its numbered version
          temp_paragraph.substitute_across_runs_with_block(/_partner_full_name_/) do |match|
            "_partner_full_name_#{partner_number}_"
          end

          temp_paragraph.substitute_across_runs_with_block(/_partner_total_quotes_/) do |match|
            "_partner_total_quotes_#{partner_number}_"
          end

          temp_paragraph.substitute_across_runs_with_block(/_parner_total_quotes_/) do |match|
            "_parner_total_quotes_#{partner_number}_"
          end

          temp_paragraph.substitute_across_runs_with_block(/_partner_sum_/) do |match|
            "_partner_sum_#{partner_number}_"
          end

          temp_paragraph.substitute_across_runs_with_block(/_%_/) do |match|
            "_%_#{partner_number}_"
          end

        end

        # Debug: show what we have in this cell after modification
        cell_text = cell.xpath('.//w:t').map(&:content).join('')
        if cell_text.include?('_partner_') || cell_text.include?('_%_')
          puts "      Cell #{cell_idx + 1} after modification: #{cell_text}"
        end
      end

      last_inserted.add_next_sibling(new_row)
      last_inserted = new_row
      puts "    ✅ Added row #{i + 1} with placeholders for partner #{partner_number}"
    end
  end
  
  # Find signature table row (Table 2)
  signature_row_index = nil
  table.rows.each_with_index do |row, row_index|
    row_text = row.cells.map { |cell|
      cell.paragraphs.map(&:to_s).join(' ')
    }.join(' ')
    
    if row_text.include?('_partner_1_full_name_') || row_text.include?('_partner_2_full_name_')
      signature_row_index = row_index
      puts "  Found signature row at index: #{signature_row_index}"
      break
    end
  end
  
  # Add rows for signature table (partners 3+)
  if signature_row_index && data['partners'].length > 2
    signature_row = table.rows[signature_row_index]
    signature_row_node = signature_row.node
    
    # Calculate how many new rows we need (2 partners per row)
    partners_to_add = data['partners'].length - 2  # Subtract the 2 already in the template
    num_signature_rows_to_add = (partners_to_add + 1) / 2  # Round up for odd numbers
    
    puts "  Adding #{num_signature_rows_to_add} new signature row(s) for #{partners_to_add} additional partner(s)..."
    
    last_inserted = signature_row_node
    num_signature_rows_to_add.times do |row_num|
      new_row = signature_row_node.dup
      
      # Calculate which partner numbers this row will handle
      first_partner_num = 3 + (row_num * 2)
      second_partner_num = first_partner_num + 1
      
      # Process cells in the new row
      cells = new_row.xpath('.//w:tc')
      cells.each_with_index do |cell, cell_idx|
        cell.xpath('.//w:p').each do |p_node|
          temp_paragraph = Docx::Elements::Containers::Paragraph.new(p_node, {}, nil)
          
          # Cell 0: Replace partner_1 with partner_3, partner_5, etc.
          # Cell 1: Replace partner_2 with partner_4, partner_6, etc.
          if cell_idx == 0
            # First cell - odd numbered partners (3, 5, 7...)
            temp_paragraph.substitute_across_runs_with_block(/_partner_1_full_name_/) do |match|
              "_partner_#{first_partner_num}_full_name_"
            end
            temp_paragraph.substitute_across_runs_with_block(/_parner_1_association_/) do |match|
              "_parner_#{first_partner_num}_association_"
            end
          elsif cell_idx == 1
            # Second cell - even numbered partners (4, 6, 8...)
            # Only add placeholders if this partner actually exists
            if second_partner_num <= data['partners'].length
              temp_paragraph.substitute_across_runs_with_block(/_partner_2_full_name_/) do |match|
                "_partner_#{second_partner_num}_full_name_"
              end
              temp_paragraph.substitute_across_runs_with_block(/_partner_2_association_/) do |match|
                "_parner_#{second_partner_num}_association_"
              end
            else
              # Clear the placeholders if partner doesn't exist
              temp_paragraph.substitute_across_runs_with_block(/_partner_2_full_name_/) do |match|
                ""
              end
              temp_paragraph.substitute_across_runs_with_block(/_partner_2_association_/) do |match|
                ""
              end
            end
          end
        end
        
        # Debug output
        cell_text = cell.xpath('.//w:t').map(&:content).join('')
        if cell_text.include?('_partner_') || cell_text.include?('_parner_')
          puts "      Signature Cell #{cell_idx + 1} after modification: #{cell_text}"
        end
      end
      
      last_inserted.add_next_sibling(new_row)
      last_inserted = new_row
      if second_partner_num <= data['partners'].length
        puts "    ✅ Added signature row #{row_num + 1} with placeholders for partners #{first_partner_num} and #{second_partner_num}"
      else
        puts "    ✅ Added signature row #{row_num + 1} with placeholder for partner #{first_partner_num}"
      end
    end
  end
end

puts "\n" + "-"*70
puts "PROCESSING TABLES - STEP 2: REPLACE TEXT"
puts "-"*70

# Second pass: Process replacements
doc.tables.each_with_index do |table, table_index|
  puts "\n📊 Processing Table #{table_index + 1}"

  table.rows.each_with_index do |row, row_index|
    row.cells.each_with_index do |cell, cell_index|
      cell.paragraphs.each_with_index do |paragraph, para_index|

        # 1. PARTNER FULL NAME
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_full_name_(?![_\w])/) do |match|
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

        # 2. PARTNER TOTAL QUOTES
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_total_quotes_(?![_\w])/) do |match|
          partners = data['partners']

          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_full_name = admin_partner ? full_name(admin_partner) : full_name(partners.first)

            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_full_name }
            result = partner_capital ? partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.') : "0"
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_total_quotes_ → #{result}"
          result
        end

        # 3. PARTNER SUM
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_sum_(?![_\w])/) do |match|
          partners = data['partners']

          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_full_name = admin_partner ? full_name(admin_partner) : full_name(partners.first)

            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_full_name }
            result = partner_capital ? "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00" : "0,00"
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_sum_ → R$ #{result}"
          result
        end

        # 4. PERCENTAGE
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_%_(?![_\w])/) do |match|
          partners = data['partners']

          if partners.length == 1
            partner_capital = data['capital']['partners'].first
            result = "#{partner_capital['percentage']}%"
          else
            admin_partner = partners.find { |p| p['is_administrator'] }
            partner_full_name = admin_partner ? full_name(admin_partner) : full_name(partners.first)

            partner_capital = data['capital']['partners'].find { |p| p['name'] == partner_full_name }
            result = partner_capital ? "#{partner_capital['percentage']}%" : "0%"
          end
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _percentage_ → #{result}"
          result
        end

        # 5. TOTAL QUOTES
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_total_quotes_(?![_\w])/) do |match|
          result = total_quotes.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _total_quotes_ → #{result}"
          result
        end

        # 6. SUM PERCENTAGE
        paragraph.substitute_across_runs_with_block(/ _sum_percentage_ /) do |match|
          total_percentage = data['capital']['partners'].sum { |p| p['percentage'] }
          result = "#{total_percentage}%"
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _sum_percentage_ → #{result}"
          result
        end

        # 7. PARTNER 1 FULL NAME - First partner (for signatures)
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_1_full_name_(?![_\w])/) do |match|
          partner = data['partners'][0]
          result = partner ? full_name(partner) : ""
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_1_full_name_ → #{result}"
          result
        end

        # 8. PARTNER 1 ASSOCIATION - First partner's role
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_parner_1_association_(?![_\w])/) do |match|
          partner = data['partners'][0]
          result = partner ? partner['association'] : ""
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _parner_1_association_ → #{result}"
          result
        end

        # 9. PARTNER 2 FULL NAME - Second partner (for signatures)
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_2_full_name_(?![_\w])/) do |match|
          partner = data['partners'][1]
          result = partner ? full_name(partner) : ""
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_2_full_name_ → #{result}"
          result
        end

        # 10. PARTNER 2 ASSOCIATION - Second partner's role
        paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_2_association_(?![_\w])/) do |match|
          partner = data['partners'][1]
          result = partner ? partner['association'] : ""
          puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_2_association_ → #{result}"
          result
        end

        # 11-16. DYNAMIC PARTNER PLACEHOLDERS (for partners 3, 4, 5, 6, etc.)
        # Process placeholders for partners 3 and beyond (for signature table)
        if data['partners'].length > 2
          (3..data['partners'].length).each do |partner_num|
            partner_index = partner_num - 1
            partner = data['partners'][partner_index]
            
            next unless partner
            
            # Partner N full name
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_#{partner_num}_full_name_(?![_\w])/) do |match|
              result = full_name(partner)
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_#{partner_num}_full_name_ → #{result}"
              result
            end
            
            # Partner N association (with typo for consistency)
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_parner_#{partner_num}_association_(?![_\w])/) do |match|
              result = partner['association']
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _parner_#{partner_num}_association_ → #{result}"
              result
            end
          end
        end

        # 17-22. NUMBERED PARTNER PLACEHOLDERS (for partner 2, 3, 4, etc.)
        # Process each additional partner's placeholders
        if data['partners'].length > 1
          # Get all partners except the admin (who is used in the first row)
          admin_partner = data['partners'].find { |p| p['is_administrator'] }
          remaining_partners = data['partners'].reject { |p| p['is_administrator'] }

          remaining_partners.each_with_index do |partner_info, idx|
            partner_num = idx + 2  # Start numbering from 2

            # Find the corresponding capital data by matching names
            partner_capital = data['capital']['partners'].find { |pc| pc['name'] == full_name(partner_info) }
            next unless partner_capital

            # Partner full name with number
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_full_name_#{partner_num}_(?![_\w])/) do |match|
              result = full_name(partner_info)
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_full_name_#{partner_num}_ → #{result}"
              result
            end

            # Partner total quotes with number
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_total_quotes_#{partner_num}_(?![_\w])/) do |match|
              result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_total_quotes_#{partner_num}_ → #{result}"
              result
            end

            # Partner total quotes with typo and number
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_parner_total_quotes_#{partner_num}_(?![_\w])/) do |match|
              result = partner_capital['quotes'].to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _parner_total_quotes_#{partner_num}_ → #{result}"
              result
            end

            # Partner sum with number
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_sum_#{partner_num}_(?![_\w])/) do |match|
              result = "#{partner_capital['value'].to_i.to_s.gsub(/\B(?=(\d{3})+(?!\d))/, '.')},00"
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _partner_sum_#{partner_num}_ → R$ #{result}"
              result
            end

            # Percentage with number (short form)
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_%_#{partner_num}_(?![_\w])/) do |match|
              result = "#{partner_capital['percentage']}%"
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _%_#{partner_num}_ → #{result}"
              result
            end

            # Percentage with number (full form)
            paragraph.substitute_across_runs_with_block(/(?<![_\w])_percentage_#{partner_num}_(?![_\w])/) do |match|
              result = "#{partner_capital['percentage']}%"
              puts "✅ Table #{table_index + 1}, Row #{row_index + 1}, Cell #{cell_index + 1}: _percentage_#{partner_num}_ → #{result}"
              result
            end
          end
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
