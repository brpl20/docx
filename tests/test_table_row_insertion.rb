#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'

puts "\n" + "="*70
puts "TABLE ROW INSERTION TEST"
puts "="*70

# Load the JSON data
data = JSON.parse(File.read('tests/test_information.json'))

puts "\n📋 Loaded data:"
puts "  Partners: #{data['partners'].length}"
data['partners'].each_with_index do |partner, idx|
  puts "    #{idx + 1}. #{partner['name']} #{partner['last_name'] || partner['last_nam']}"
end

# Open the template
template_file = data['partners'].length == 1 ? 'tests/CS-TEMPLATE-INDIVIDUAL.docx' : 'tests/CS-TEMPLATE.docx'
doc = Docx::Document.open(template_file)

puts "\n📊 Processing Tables..."

# Process each table
doc.tables.each_with_index do |table, table_index|
  puts "\nTable #{table_index + 1}:"
  puts "  Rows: #{table.row_count}"
  puts "  Columns: #{table.column_count}"
  
  # Find the partner row (row that contains partner placeholders)
  partner_row_index = nil
  
  table.rows.each_with_index do |row, row_index|
    # Get all text from the row
    row_text = row.cells.map { |cell| 
      cell.paragraphs.map(&:to_s).join(' ')
    }.join(' ')
    
    # Check if this row contains partner-specific placeholders
    if row_text.include?('_partner_full_name_') || 
       row_text.include?('_partner_total_quotes_') || 
       row_text.include?('_partner_sum_')
      
      partner_row_index = row_index
      puts "  Found partner row at index: #{partner_row_index}"
      puts "  Row content preview: #{row_text[0..100]}..."
      break
    end
  end
  
  # If we found a partner row and have multiple partners, insert additional rows
  if partner_row_index && data['partners'].length > 1
    puts "\n  📝 Inserting rows for multiple partners..."
    
    # Get the partner row
    partner_row = table.rows[partner_row_index]
    partner_row_node = partner_row.node
    
    # We need to insert rows for all partners except the first one
    # (the first one will use the existing row)
    num_partners = data['partners'].length
    num_rows_to_add = num_partners - 1
    
    puts "  Will add #{num_rows_to_add} new row(s)"
    
    # Insert new rows after the partner row
    last_inserted = partner_row_node
    
    num_rows_to_add.times do |i|
      puts "  Adding row #{i + 1}..."
      
      # Clone the partner row
      new_row = partner_row_node.dup
      
      # Insert after the last inserted row
      last_inserted.add_next_sibling(new_row)
      last_inserted = new_row
      
      puts "    ✅ Row added successfully"
    end
    
    puts "\n  Table now has #{table.row_count} rows"
  else
    if !partner_row_index
      puts "  ℹ️ No partner row found in this table"
    else
      puts "  ℹ️ Only one partner, no rows to add"
    end
  end
end

# Save the output
output_path = 'tests/CS-TEMPLATE-table-rows-added.docx'
doc.save(output_path)

puts "\n" + "="*70
puts "COMPLETE"
puts "="*70
puts "\n📄 Output saved to: #{output_path}"
puts "✨ Table row insertion completed successfully!"