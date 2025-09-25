# docx-br

[![Gem Version](https://badge.fury.io/rb/docx-br.svg)](https://badge.fury.io/rb/docx-br)
[![Ruby](https://github.com/ruby-docx/docx/workflows/Ruby/badge.svg)](https://github.com/ruby-docx/docx/actions?query=workflow%3ARuby)
[![Coverage Status](https://coveralls.io/repos/github/ruby-docx/docx/badge.svg?branch=master)](https://coveralls.io/github/ruby-docx/docx?branch=master)
[![Gitter](https://badges.gitter.im/ruby-docx/community.svg)](https://gitter.im/ruby-docx/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

A ruby library/gem for interacting with `.docx` files. This is a fork of the original `docx` gem with improvements to handle text substitution across fragmented text runs - a common issue when Word internally splits placeholders across multiple XML nodes.

**Key improvement:** Fixes the well-known issue where placeholders like `{{placeholder}}` or `_placeholder_` fail to be replaced because Word fragments them across multiple text runs.

## Usage

### Prerequisites

- Ruby 2.6 or later

### Install

Add the following line to your application's Gemfile:

```ruby
gem 'docx-br'
```

And then execute:

```shell
bundle install
```

Or install it yourself as:

```shell
gem install docx-br
```

### Reading

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('example.docx')

# Retrieve and display paragraphs
doc.paragraphs.each do |p|
  puts p
end

# Retrieve and display bookmarks, returned as hash with bookmark names as keys and objects as values
doc.bookmarks.each_pair do |bookmark_name, bookmark_object|
  puts bookmark_name
end
```

Don't have a local file but a buffer? Docx handles those too:

```ruby
require 'docx'

# Create a Docx::Document object from a remote file
doc = Docx::Document.open(buffer)

# Everything about reading is the same as shown above
```

### Rendering html
``` ruby
require 'docx'

# Retrieve and display paragraphs as html
doc = Docx::Document.open('example.docx')
doc.paragraphs.each do |p|
  puts p.to_html
end
```

### Reading tables

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('tables.docx')

first_table = doc.tables[0]
puts first_table.row_count
puts first_table.column_count
puts first_table.rows[0].cells[0].text
puts first_table.columns[0].cells[0].text

# Iterate through tables
doc.tables.each do |table|
  table.rows.each do |row| # Row-based iteration
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table.columns.each do |column| # Column-based iteration
    column.cells.each do |cell|
      puts cell.text
    end
  end
end
```

### Writing

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('example.docx')

# Insert a single line of text after one of our bookmarks
doc.bookmarks['example_bookmark'].insert_text_after("Hello world.")

# Insert multiple lines of text at our bookmark
doc.bookmarks['example_bookmark_2'].insert_multiple_lines_after(['Hello', 'World', 'foo'])

# Remove paragraphs
doc.paragraphs.each do |p|
  p.remove! if p.to_s =~ /TODO/
end

# Substitute text, preserving formatting
# Method 1: Traditional approach (only works if placeholder is not fragmented)
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute('_placeholder_', 'replacement value')
  end
end

# Method 2: NEW - Substitution across fragmented runs (handles Word's text fragmentation)
doc.paragraphs.each do |p|
  # This will find and replace placeholders even if Word split them across multiple runs
  p.substitute_across_runs('{{placeholder}}', 'replacement value')
  p.substitute_across_runs('_placeholder_', 'another value')
  
  # Also works with regex patterns
  p.substitute_across_runs(/\{\{(\w+)\}\}/, 'replaced')
end

# Method 3: NEW - Consolidate fragmented runs before substitution
doc.paragraphs.each do |p|
  # Merges adjacent runs with identical formatting
  p.consolidate_text_runs
  
  # Now traditional substitution is more likely to work
  p.each_text_run do |tr|
    tr.substitute('_placeholder_', 'replacement value')
  end
end

# Substitute text with access to captures, note block arg is a MatchData, a bit
# different than String.gsub. https://ruby-doc.org/3.3.7/MatchData.html
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute_with_block(/total: (\d+)/) { |match_data| "total: #{match_data[1].to_i * 10}" }
  end
end

# NEW - Block substitution across fragmented runs
doc.paragraphs.each do |p|
  p.substitute_across_runs_with_block(/total: (\d+)/) { |match_data| 
    "total: #{match_data[1].to_i * 10}" 
  }
end

# NEWEST - Advanced regex substitution with automatic word boundary protection
doc.paragraphs.each do |p|
  # Simple string patterns get automatic word boundary protection
  p.substitute_across_runs_with_block_regex("_company_name_") do |match|
    "ACME Corporation"
  end
  
  # Dynamic patterns work seamlessly
  partner_number = 2
  p.substitute_across_runs_with_block_regex("_partner_#{partner_number}_name_") do |match|
    "John Smith"
  end
  
  # Custom regex patterns still supported when needed
  p.substitute_across_runs_with_block_regex(/(?<![_\w])_office_(\w+)_(?![_\w])/) do |match|
    office_data[match[1]]  # Uses captured group
  end
end

# Save document to specified path
doc.save('example-edited.docx')
```

### Writing to tables

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('tables.docx')

# Iterate over each table
doc.tables.each do |table|
  last_row = table.rows.last

  # Copy last row and insert a new one before last row
  new_row = last_row.copy
  new_row.insert_before(last_row)

  # Substitute text in each cell of this new row
  new_row.cells.each do |cell|
    cell.paragraphs.each do |paragraph|
      paragraph.each_text_run do |text|
        text.substitute('_placeholder_', 'replacement value')
      end
    end
  end
end

doc.save('tables-edited.docx')
```

### Advanced

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')

# The Nokogiri::XML::Node on which an element is based can be accessed using #node
d.paragraphs.each do |p|
  puts p.node.inspect
end

# The #xpath and #at_xpath methods are delegated to the node from the element, saving a step
p_element = d.paragraphs.first
p_children = p_element.xpath("//child::*") # selects all children
p_child = p_element.at_xpath("//child::*") # selects first child
```

### Writing and Manipulating Styles
``` ruby
require 'docx'

d = Docx::Document.open('example.docx')
existing_style = d.styles_configuration.style_of("Heading 1")
existing_style.font_color = "000000"

# see attributes below
new_style = d.styles_configuration.add_style("Red", name: "Red", font_color: "FF0000", font_size: 20)
new_style.bold = true

d.paragraphs.each do |p|
  p.style = "Red"
end

d.paragraphs.each do |p|
  p.style = "Heading 1"
end

d.styles_configuration.remove_style("Red")
```

#### Style Attributes

The following is a list of attributes and what they control within the style.

- **id**: The unique identifier of the style. (required)
- **name**: The human-readable name of the style. (required)
- **type**: Indicates the type of the style (e.g., paragraph, character).
- **keep_next**: Boolean value controlling whether to keep a paragraph and the next one on the same page. Valid values: `true`/`false`.
- **keep_lines**: Boolean value specifying whether to keep all lines of a paragraph together on one page. Valid values: `true`/`false`.
- **page_break_before**: Boolean value indicating whether to insert a page break before the paragraph. Valid values: `true`/`false`.
- **widow_control**: Boolean value controlling widow and orphan lines in a paragraph. Valid values: `true`/`false`.
- **shading_style**: Defines the shading pattern style.
- **shading_color**: Specifies the color of the shading pattern. Valid values: Hex color codes.
-  **shading_fill**: Indicates the background fill color of shading.
-  **suppress_auto_hyphens**: Boolean value controlling automatic hyphenation. Valid values: `true`/`false`.
-  **bidirectional_text**: Boolean value indicating if the paragraph contains bidirectional text. Valid values: `true`/`false`.
-  **spacing_before**: Defines the spacing before a paragraph.
-  **spacing_after**: Specifies the spacing after a paragraph.
-  **line_spacing**: Indicates the line spacing of a paragraph.
-  **line_rule**: Defines how line spacing is calculated.
-  **indent_left**: Sets the left indentation of a paragraph.
-  **indent_right**: Specifies the right indentation of a paragraph.
-  **indent_first_line**: Indicates the first line indentation of a paragraph.
-  **align**: Controls the text alignment within a paragraph.
-  **font**: Sets the font for different scripts (ASCII, complex script, East Asian, etc.).
-  **font_ascii**: Specifies the font for ASCII characters.
-  **font_cs**: Indicates the font for complex script characters.
-  **font_hAnsi**: Sets the font for high ANSI characters.
-  **font_eastAsia**: Specifies the font for East Asian characters.
-  **bold**: Boolean value controlling bold formatting. Valid values: `true`/`false`.
-  **italic**: Boolean value indicating italic formatting. Valid values: `true`/`false`.
-  **caps**: Boolean value controlling capitalization. Valid values: `true`/`false`.
-  **small_caps**: Boolean value specifying small capital letters. Valid values: `true`/`false`.
-  **strike**: Boolean value indicating strikethrough formatting. Valid values: `true`/`false`.
-  **double_strike**: Boolean value defining double strikethrough formatting. Valid values: `true`/`false`.
-  **outline**: Boolean value specifying outline effects. Valid values: `true`/`false`.
-  **outline_level**: Indicates the outline level in a document's hierarchy.
-  **font_color**: Sets the text color. Valid values: Hex color codes.
-  **font_size**: Controls the font size.
-  **font_size_cs**: Specifies the font size for complex script characters.
-  **underline_style**: Indicates the style of underlining.
-  **underline_color**: Specifies the color of the underline. Valid values: Hex color codes.
-  **spacing**: Controls character spacing.
-  **kerning**: Sets the space between characters.
-  **position**: Controls the position of characters (superscript/subscript).
-  **text_fill_color**: Sets the fill color of text. Valid values: Hex color codes.
-  **vertical_alignment**: Controls the vertical alignment of text within a line.
-  **lang**: Specifies the language tag for the text.

## Text Substitution Improvements

### The Problem

Word internally fragments text across multiple `<w:r>` (text run) XML nodes, even for text that appears continuous in the document. This causes standard find/replace operations to fail when searching for placeholders like `{{name}}` or `_placeholder_`.

For example, the text `{{name}}` might be stored in the XML as:
```xml
<w:r><w:t>{{</w:t></w:r>
<w:r><w:t>na</w:t></w:r>
<w:r><w:t>me</w:t></w:r>
<w:r><w:t>}}</w:t></w:r>
```

### The Solution

This gem provides new methods to handle substitution across fragmented runs:

1. **`substitute_across_runs(pattern, replacement)`** - Searches for and replaces text across all runs in a paragraph
2. **`substitute_across_runs_with_block(pattern, &block)`** - Same as above but with block support for dynamic replacements  
3. **`substitute_across_runs_with_block_regex(pattern, &block)`** ⭐ **NEW!** - Advanced substitution with automatic word boundary protection
4. **`consolidate_text_runs`** - Merges adjacent runs with identical formatting to reduce fragmentation

### Example Usage

```ruby
require 'docx'

doc = Docx::Document.open('template.docx')

doc.paragraphs.each do |paragraph|
  # Replace even if {{client_name}} is fragmented
  paragraph.substitute_across_runs('{{client_name}}', 'ABC Corporation')
  
  # Handle multiple placeholders
  paragraph.substitute_across_runs('{{date}}', Date.today.to_s)
  paragraph.substitute_across_runs('{{amount}}', '$1,000')
  
  # Use regex with captures
  paragraph.substitute_across_runs_with_block(/\{\{price_(\d+)\}\}/) do |match|
    price = match[1].to_i * 1.1  # Add 10% markup
    "$#{price}"
  end
end

doc.save('output.docx')
```

### Advanced Regex Substitution ⭐ NEW!

The new `substitute_across_runs_with_block_regex` method provides automatic word boundary protection and prevents placeholder overlapping issues:

```ruby
require 'docx'

doc = Docx::Document.open('template.docx')

doc.paragraphs.each do |paragraph|
  # Simple usage - automatic word boundary protection
  # Pattern "_company_name_" becomes /(?<![_\w])_company_name_(?![_\w])/
  paragraph.substitute_across_runs_with_block_regex("_company_name_") do |match|
    "ACME Corporation"
  end
  
  # Dynamic patterns work seamlessly
  partner_number = 2
  paragraph.substitute_across_runs_with_block_regex("_partner_#{partner_number}_name_") do |match|
    partner_data[partner_number]['name']
  end
  
  # Conditional replacements
  paragraph.substitute_across_runs_with_block_regex("_optional_clause_") do |match|
    include_clause ? "This clause is included." : ""
  end
  
  # Custom complex regex patterns when needed
  paragraph.substitute_across_runs_with_block_regex(/(?<![_\w])_office_(\w+)_(?![_\w])/) do |match|
    field_name = match[1]  # captured group
    office_data[field_name] || "N/A"
  end
end

doc.save('advanced_output.docx')
```

**Key Benefits:**
- 🛡️ **Automatic Protection**: String patterns get word boundary protection automatically
- 🧹 **Cleaner Code**: No need to write complex regex patterns manually  
- 🔒 **Overlap Prevention**: Prevents issues like `_partner_name_2_` matching `_partner_name_`
- 🎯 **Flexible**: Still supports custom regex when needed
- ⚡ **Reliable**: Handles Word's text fragmentation seamlessly

**Migration Guide:**
```ruby
# OLD - Manual regex (error-prone)
paragraph.substitute_across_runs_with_block(/(?<![_\w])_field_(?![_\w])/) { value }

# NEW - Clean and safe  
paragraph.substitute_across_runs_with_block_regex("_field_") { value }
```

## Placeholder Debugger

The placeholder debugger is a powerful tool to help you identify and fix placeholder replacement issues in your Word documents. It validates that placeholders can be found and replaced correctly, even when Word fragments them across multiple XML nodes.

### Features

- 🔍 **Automatic Detection**: Finds placeholders in your document
- ⚠️ **Fragmentation Detection**: Identifies when Word has split placeholders
- ✅ **Validation**: Tests if replacements will work
- 🏗️ **Code Generation**: Creates a custom replacer class for your template
- 📊 **Detailed Reporting**: Shows success/failure for each placeholder

### Quick Start

```ruby
require 'docx'

# Quick check with underline placeholders
Docx::Debugger.quick_check(
  'template.docx',
  :underline,
  ['office_name', 'partner_name', 'date']
)
```

### Supported Placeholder Formats

| Format | Pattern | Example |
|--------|---------|---------|
| `:underline` | `_text_` | `_office_name_` |
| `:mustache` | `{ text }` | `{ office_name }` |
| `:double_mustache` | `{{ text }}` | `{{ office_name }}` |
| `:angle` | `< text >` | `< office_name >` |
| `:square` | `[ text ]` | `[ office_name ]` |
| `:dollar` | `${text}` | `${office_name}` |

### Basic Debugging

```ruby
debugger = Docx::Debugger.analyze('template.docx') do |config|
  config.placeholder_type = :double_mustache
  config.set_placeholders(['client_name', 'contract_date', 'amount'])
end

results = debugger.debug!
```

### Generated Replacer Class

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

## Replacement Validation

The gem includes a comprehensive validation system to ensure all placeholders are properly replaced:

```ruby
# Validate that all replacements worked
validator = Docx::ReplacementValidator.validate(
  'original_template.docx',
  'processed_output.docx'
)

validator.report  # Shows detailed validation results

if validator.passed?
  puts "All placeholders successfully replaced!"
else
  puts "Some placeholders were missed:"
  validator.failed_placeholders.each do |failed|
    puts "- #{failed[:placeholder]} in paragraph #{failed[:paragraph_index]}"
  end
end
```

## Development

### TODO

* Calculate element formatting based on values present in element properties as well as properties inherited from parents
* Default formatting of inserted elements to inherited values
* Implement formattable elements
* Easier multi-line text insertion at a single bookmark (inserting paragraph nodes after the one containing the bookmark)
* **Debugger with output class integration** - Enhanced debugger that works with generated replacer classes
* **Advanced validation features** - More sophisticated error detection and remediation

