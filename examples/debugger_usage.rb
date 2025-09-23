#!/usr/bin/env ruby
# frozen_string_literal: true

require 'docx'

# Example 1: Quick debugging with underline placeholders
puts "\n=== EXAMPLE 1: Underline Placeholders ==="
puts "Checking template with underline placeholders like _office_name_"

Docx::Debugger.quick_check(
  'template.docx',
  :underline,
  ['office_name', 'partner_qualification', 'date', 'client_name']
)

# Example 2: Debugging with mustache-style placeholders
puts "\n=== EXAMPLE 2: Mustache Placeholders ==="
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :mustache
  config.add_placeholder('office_name')
  config.add_placeholder('partner_qualification')
  config.add_placeholder('client_name')
end

results = debugger.debug!

# Example 3: Double mustache (Handlebars) style
puts "\n=== EXAMPLE 3: Double Mustache Placeholders ==="
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders(['first_name', 'last_name', 'company', 'amount'])
end

results = debugger.debug!

# Test with actual replacement and save
debugger.test_replacement(save_as: 'test_output.docx')

# Example 4: Custom pattern
puts "\n=== EXAMPLE 4: Custom Pattern ==="
debugger = Docx::Debugger.analyze('template.docx') do |config|
  # Custom pattern for placeholders like %PLACEHOLDER%
  config.custom_pattern = /%(\w+)%/
  config.add_placeholder('CUSTOM_VAR')
  config.add_placeholder('ANOTHER_VAR')
end

debugger.debug!

# Example 5: Full workflow with generated replacer class
puts "\n=== EXAMPLE 5: Complete Workflow ==="

# Step 1: Debug the template
debugger = Docx::Debugger.analyze('contract_template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders([
    'company_name',
    'contract_date',
    'contract_amount',
    'payment_terms',
    'signatory_name',
    'signatory_title'
  ])
end

results = debugger.debug!

# Step 2: If successful, a replacer class will be generated
# The file will be created in the current directory

# Step 3: Use the generated replacer class
if results[:failed] == 0
  puts "\n📝 Using the generated replacer class..."
  
  # The generated class would be used like this:
  # require_relative 'double_mustache_replacer'
  # 
  # replacer = DoubleMustacheReplacer.new('contract_template.docx')
  # replacer.company_name = 'ABC Corporation'
  # replacer.contract_date = Date.today.to_s
  # replacer.contract_amount = '$50,000'
  # replacer.payment_terms = 'Net 30 days'
  # replacer.signatory_name = 'John Doe'
  # replacer.signatory_title = 'CEO'
  # 
  # replacer.process!('signed_contract.docx')
end

# Example 6: Debugging tables
puts "\n=== EXAMPLE 6: Placeholders in Tables ==="
debugger = Docx::Debugger.analyze('invoice_template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders([
    'invoice_number',
    'item_description',
    'quantity',
    'unit_price',
    'total'
  ])
end

debugger.debug!

# Example 7: Angle bracket placeholders
puts "\n=== EXAMPLE 7: Angle Bracket Placeholders ==="
Docx::Debugger.quick_check(
  'template.docx',
  :angle,
  ['variable1', 'variable2', 'variable3']
)

# Example 8: Programmatic access to results
puts "\n=== EXAMPLE 8: Processing Results Programmatically ==="
debugger = Docx::Debugger.analyze('report_template.docx') do |config|
  config.placeholder_type = :dollar  # ${placeholder} style
  config.set_placeholders(['title', 'author', 'date', 'content'])
end

results = debugger.debug!

# Process the results
if results[:failed] > 0
  puts "\n⚠️  Failed placeholders need attention:"
  results[:details].each do |detail|
    if !detail[:success]
      puts "  - #{detail[:placeholder][:formatted]}: #{detail[:error]}"
    end
  end
else
  puts "\n✅ All placeholders validated successfully!"
  puts "You can now safely use the generated replacer class."
end