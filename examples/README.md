# Docx-BR Examples

This folder contains practical examples showing how to use the docx-br gem's advanced features.

## Examples

### 1. Simple Debugger (`simple_debugger.rb`)

The most basic way to debug placeholders in your documents:

```ruby
# One-liner check
Docx::Debugger.quick_check('template.docx', :double_mustache, ['name', 'date'])

# Basic configuration
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :underline
  config.set_placeholders(['office_name', 'partner_name'])
end

results = debugger.debug!
```

**Use this when**: You want to quickly check if your placeholders will work.

### 2. Debugger with Generated Class (`debugger_with_generated_class.rb`)

Shows how to use the debugger and then utilize the generated replacer class:

```ruby
# Debug first
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders(['client_name', 'amount'])
end

results = debugger.debug!

# Then use the generated class
if results[:failed] == 0
  require_relative 'double_mustache_replacer'
  
  replacer = DoubleMustacheReplacer.new('template.docx')
  replacer.client_name = 'ABC Corp'
  replacer.amount = '$50,000'
  replacer.process!('output.docx')
end
```

**Use this when**: You want to generate a custom replacer class for your specific template.

**Note**: This example includes TODO items for enhanced integration features planned for future releases.

### 3. Replacement Checker (`replacement_checker.rb`)

Validates that all placeholders were properly replaced in the final document:

```ruby
# Validate replacements
validator = Docx::ReplacementValidator.validate(
  'original_template.docx',
  'processed_output.docx'
)

validator.report

if validator.passed?
  puts "✅ All placeholders replaced successfully!"
else
  puts "❌ Some placeholders were missed"
  # Show detailed error information
end
```

**Use this when**: You need to ensure 100% reliability in production document generation.

### 4. Advanced Regex Substitution (`substitute_across_runs_with_block_regex_usage.rb`)

Demonstrates the new advanced substitution method with automatic word boundary protection:

```ruby
# Simple usage with automatic word boundaries
paragraph.substitute_across_runs_with_block_regex("_company_name_") do |match|
  "ACME Corporation Ltd."
end

# Dynamic patterns
partner_number = 2
paragraph.substitute_across_runs_with_block_regex("_partner_#{partner_number}_name_") do |match|
  "John Smith"
end

# Custom complex regex when needed
paragraph.substitute_across_runs_with_block_regex(/(?<![_\w])_office_\w+_(?![_\w])/) do |match|
  "Various office information"
end
```

**Use this when**: You need safe, reliable placeholder replacement that prevents overlapping issues.

### 5. Method Comparison (`method_comparison.rb`)

Shows the evolution from basic to advanced substitution methods:

```ruby
# OLD - Manual regex (error-prone)
paragraph.substitute_across_runs_with_block(/(?<![_\w])_field_(?![_\w])/) { value }

# NEW - Clean and safe
paragraph.substitute_across_runs_with_block_regex("_field_") { value }
```

**Use this when**: You're migrating from older substitution methods or learning the differences.

## Running the Examples

```bash
# Make sure you have the gem loaded
ruby -I lib examples/simple_debugger.rb

# Or if you have the gem installed
ruby examples/simple_debugger.rb
```

## Example Templates

For these examples to work, you'll need Word documents with placeholders like:

- `{{ client_name }}` (double mustache)
- `_office_name_` (underline)
- `< partner_name >` (angle brackets)
- `${amount}` (dollar style)

## Production Workflow

A typical production workflow would be:

1. **Debug** your template with `simple_debugger.rb`
2. **Generate** a custom replacer class 
3. **Use** the generated class in your application
4. **Validate** the output with `replacement_checker.rb`

This ensures reliable, error-free document generation! 🚀