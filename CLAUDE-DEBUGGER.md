# Claude Debugging Guide for DOCX Processing

This document outlines debugging techniques used to troubleshoot DOCX template replacement issues in Ruby scripts.

## Common Debugging Scenarios

### 1. Partner Name Duplication Issue

**Problem**: Same partner name appearing in multiple table rows instead of different partners.

**Debug Strategy**: 
- Add debug output to see which partner data is being selected for each iteration
- Check the partner matching logic between `data['partners']` and `data['capital']['partners']` arrays

**Example Debug Code**:
```ruby
partner_full_name = full_name(partner_info)
partner_capital = data['capital']['partners'].find { |pc| pc['name'] == partner_full_name }
puts "🔍 DEBUG: Looking for partner #{partner_full_name}, found: #{partner_capital ? 'YES' : 'NO'}"
```

### 2. Template Placeholder Format Issues

**Problem**: Placeholders not being found or replaced (e.g., `_%_` vs `_% _`).

**Debug Strategy**:
- Inspect original template content before any modifications
- Compare expected placeholder format with actual template content

**Example Debug Code**:
```ruby
# Show all cell content in template tables
cells.each_with_index do |cell, cell_idx|
  original_text = cell.xpath('.//w:t').map(&:content).join('')
  puts "📋 Cell #{cell_idx + 1} original content: '#{original_text}'"
end
```

### 3. Table Row Creation Verification

**Problem**: New table rows not being created with correct numbered placeholders.

**Debug Strategy**:
- Show modified cell content after placeholder conversion
- Verify numbered placeholders are being created correctly

**Example Debug Code**:
```ruby
# After placeholder modifications
cell_text = cell.xpath('.//w:t').map(&:content).join('')
if cell_text.include?('_partner_') || cell_text.include?('_%_')
  puts "Cell #{cell_idx + 1} after modification: #{cell_text}"
end
```

### 4. Regex Pattern Debugging

**Problem**: Replacement patterns not matching expected text.

**Debug Strategy**:
- Add debug output within replacement blocks
- Show before/after text for each replacement

**Example Debug Code**:
```ruby
paragraph.substitute_across_runs_with_block(/(?<![_\w])_partner_full_name_#{partner_num}_(?![_\w])/) do |match|
  result = full_name(partner_info)
  puts "DEBUGGER: Replacing _partner_full_name_#{partner_num}_ → #{result}"
  result
end
```

## Key Debugging Patterns

### 1. JSON Data Structure Validation
```ruby
# Verify data structure
puts "Partners count: #{data['partners'].length}"
data['partners'].each_with_index do |p, idx|
  puts "Partner #{idx}: #{full_name(p)}"
end
data['capital']['partners'].each_with_index do |pc, idx|
  puts "Capital #{idx}: #{pc['name']}"
end
```

### 2. Template Content Inspection
```ruby
# Check what's actually in the template
doc.tables.each_with_index do |table, table_idx|
  puts "Table #{table_idx + 1}:"
  table.rows.each_with_index do |row, row_idx|
    puts "  Row #{row_idx + 1}:"
    row.cells.each_with_index do |cell, cell_idx|
      cell_text = cell.paragraphs.map(&:to_s).join(' ')
      puts "    Cell #{cell_idx + 1}: '#{cell_text}'"
    end
  end
end
```

### 3. Replacement Result Verification
```ruby
# Track all replacements
puts "✅ Replaced #{placeholder} → #{result} (Table #{table_idx}, Row #{row_idx}, Cell #{cell_idx})"
```

## Common Issues Found

1. **JSON Structure Mismatches**: `"last_nam"` vs `"last_name"` typos
2. **Placeholder Format Variations**: `_%_` vs `_% _` (with space)
3. **Array Index vs Name Matching**: Using wrong logic to match partner data
4. **Admin Partner Logic**: Not accounting for administrator preference in multi-partner scenarios

## Debugging Cleanup

Always remove debug output before committing:
```ruby
# Remove lines like:
puts "🔍 DEBUG: ..."
puts "DEBUGGER: ..."
puts "📋 Cell content: ..."
```

## Git Workflow for Debugging Sessions

1. Create temporary debug commits during investigation
2. Use `git rebase -i` to squash debug commits before final push
3. Ensure final commit follows Conventional Commits format
4. Document findings in commit message or separate documentation