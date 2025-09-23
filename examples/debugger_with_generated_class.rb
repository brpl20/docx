#!/usr/bin/env ruby
# frozen_string_literal: true

# Debugger with Generated Class Example
# Shows how to use the debugger and then use the generated replacer class
# TODO: This needs to be updated to integrate better with output classes

require 'docx'

puts "Debugger with Generated Class Example"
puts "===================================="

template_path = 'contract_template.docx'

# Step 1: Debug the template
puts "\n📋 Step 1: Debugging template..."
debugger = Docx::Debugger.analyze(template_path) do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders([
    'client_name',
    'contract_date', 
    'amount',
    'payment_terms'
  ])
end

results = debugger.debug!

# Step 2: Use the generated class if debugging was successful
if results[:failed] == 0
  puts "\n🎉 All placeholders validated! Using generated replacer..."
  
  # The debugger generates a file like 'double_mustache_replacer.rb'
  # In a real scenario, you would require it:
  # require_relative 'double_mustache_replacer'
  
  puts "\n📝 Generated class usage:"
  puts <<~USAGE
    # This is how you would use the generated class:
    
    require_relative 'double_mustache_replacer'
    
    replacer = DoubleMustacheReplacer.new('#{template_path}')
    replacer.client_name = 'ABC Corporation'
    replacer.contract_date = '#{Date.today}'
    replacer.amount = '$50,000'
    replacer.payment_terms = 'Net 30 days'
    
    # Check if all required fields are set
    if replacer.ready?
      replacer.process!('final_contract.docx')
      puts "✅ Contract generated successfully!"
    else
      puts "❌ Missing: \#{replacer.missing_placeholders.join(', ')}"
    end
  USAGE
  
else
  puts "\n❌ Template validation failed."
  puts "Please fix the issues shown above before generating documents."
  
  # Show which placeholders failed
  failed_placeholders = results[:details]
    .select { |d| !d[:success] }
    .map { |d| d[:placeholder][:formatted] }
  
  puts "\nFailed placeholders:"
  failed_placeholders.each { |p| puts "  - #{p}" }
end

puts "\n" + "="*50
puts "💡 TODO: Enhanced Integration"
puts "="*50
puts <<~TODO
This example shows the current workflow, but we plan to enhance it with:

1. **Automatic Class Loading**: The debugger could automatically 
   require and instantiate the generated class.

2. **Interactive Setup**: A guided process to set placeholder values.

3. **Validation Integration**: Real-time validation as you set values.

4. **Template Variants**: Support for multiple template versions.

5. **Batch Processing**: Process multiple documents with the same template.

These features are on our TODO list for future releases!
TODO

puts "Done! 🚀"