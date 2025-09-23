# Docx-BR Placeholder Debugger

The placeholder debugger is a powerful tool to help you identify and fix placeholder replacement issues in your Word documents. It validates that placeholders can be found and replaced correctly, even when Word fragments them across multiple XML nodes.

## Features

- 🔍 **Automatic Detection**: Finds placeholders in your document
- ⚠️ **Fragmentation Detection**: Identifies when Word has split placeholders
- ✅ **Validation**: Tests if replacements will work
- 🏗️ **Code Generation**: Creates a custom replacer class for your template
- 📊 **Detailed Reporting**: Shows success/failure for each placeholder

## Quick Start

```ruby
require 'docx'

# Quick check with underline placeholders
Docx::Debugger.quick_check(
  'template.docx',
  :underline,
  ['office_name', 'partner_name', 'date']
)
```

## Supported Placeholder Formats

The debugger supports multiple placeholder formats:

| Format | Pattern | Example |
|--------|---------|---------|
| `:underline` | `_text_` | `_office_name_` |
| `:mustache` | `{ text }` | `{ office_name }` |
| `:double_mustache` | `{{ text }}` | `{{ office_name }}` |
| `:angle` | `< text >` | `< office_name >` |
| `:square` | `[ text ]` | `[ office_name ]` |
| `:dollar` | `${text}` | `${office_name}` |

## Detailed Usage

### Basic Debugging

```ruby
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders(['client_name', 'contract_date', 'amount'])
end

results = debugger.debug!
```

### Output Example

```
============================================================
DOCX PLACEHOLDER DEBUGGER
============================================================

Template: template.docx
Pattern Type: double_mustache
Pattern: {{ placeholder_name }}
------------------------------------------------------------

📊 VALIDATION RESULTS:
------------------------------------------------------------

1. ✅ {{ client_name }}
   Found: 3 occurrence(s)
   Paragraphs: 1, 5, 12
   Test replacement: SUCCESS

2. ❌ {{ contract_date }}
   Error: Placeholder found but fragmented across text runs. Use substitute_across_runs method.
   Found: 1 occurrence(s)
   Paragraphs: 3

3. ✅ {{ amount }}
   Found: 2 occurrence(s)
   Paragraphs: 8, 15
   Test replacement: SUCCESS

============================================================
SUMMARY:
  Total placeholders: 3
  ✅ Successful: 2
  ❌ Failed: 1
  Success Rate: 66.67%
============================================================
```

### Custom Patterns

You can define custom regex patterns for unique placeholder formats:

```ruby
debugger = Docx::Debugger.analyze('template.docx') do |config|
  # Custom pattern for %PLACEHOLDER%
  config.custom_pattern = /%(\w+)%/
  config.add_placeholder('CUSTOM_VAR')
end

debugger.debug!
```

### Test Replacements

Generate a test document with random values to verify replacements:

```ruby
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders(['name', 'date', 'signature'])
end

# Creates a test document with placeholders replaced by random values
debugger.test_replacement(save_as: 'test_output.docx')
```

## Generated Replacer Class

When all placeholders validate successfully, the debugger generates a custom replacer class:

```ruby
# Generated file: double_mustache_replacer.rb

class DoubleMustacheReplacer
  def initialize(template_path)
    @document = Docx::Document.open(template_path)
    # ...
  end
  
  def client_name=(value)
    @replacements[:client_name] = value
  end
  
  def contract_date=(value)
    @replacements[:contract_date] = value
  end
  
  def replace_all!
    # Replaces all placeholders using substitute_across_runs
  end
  
  def process!(output_path)
    replace_all!
    save(output_path)
  end
end
```

### Using the Generated Class

```ruby
require_relative 'double_mustache_replacer'

replacer = DoubleMustacheReplacer.new('template.docx')
replacer.client_name = 'ABC Corporation'
replacer.contract_date = '2024-01-15'
replacer.amount = '$50,000'

replacer.process!('final_contract.docx')
```

## Complete Workflow Example

```ruby
# Step 1: Debug your template
debugger = Docx::Debugger.analyze('contract.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders([
    'company_name',
    'contract_date',
    'amount',
    'terms'
  ])
end

results = debugger.debug!

# Step 2: If successful, use the generated replacer
if results[:failed] == 0
  require_relative 'double_mustache_replacer'
  
  replacer = DoubleMustacheReplacer.new('contract.docx')
  replacer.company_name = 'Tech Corp'
  replacer.contract_date = Date.today
  replacer.amount = '$100,000'
  replacer.terms = 'Net 30'
  
  replacer.process!('signed_contract.docx')
end
```

## Handling Fragmented Placeholders

If the debugger reports that placeholders are fragmented:

1. The debugger will detect this and warn you
2. The generated replacer class automatically uses `substitute_across_runs`
3. This ensures replacements work even with fragmented text

## Tips

1. **Start Simple**: Use the `quick_check` method for rapid testing
2. **Check Fragmentation**: If replacements fail, the debugger will tell you why
3. **Use Generated Classes**: They handle all the complexity for you
4. **Test First**: Use `test_replacement` to verify before production use

## API Reference

### Docx::Debugger.analyze(template_path)

Creates a debugger instance for detailed analysis.

### Docx::Debugger.quick_check(template_path, type, placeholders)

Quick validation of placeholders.

### Configuration Options

- `placeholder_type`: Symbol for placeholder format
- `placeholders`: Array of placeholder names
- `custom_pattern`: Regex for custom formats

### Results Structure

```ruby
{
  total: 5,
  successful: 4,
  failed: 1,
  details: [
    {
      placeholder: { name: 'client_name', formatted: '{{ client_name }}' },
      success: true,
      found_count: 2,
      paragraph_indices: [1, 5]
    },
    # ...
  ]
}
```