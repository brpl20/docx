#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('lib', __dir__))
require 'docx'
require 'docx/debugger'

# Define ALL placeholders including duplicates
# This represents the exact order and frequency they appear in your document
placeholders = %w[
  office_name
  partner_qualification
  office_name
  office_city
  office_state
  office_address
  office_zip_code
  office_total_value
  office_quotes
  office_quote_value
  partner_full_name
  partner_total_quotes
  office_quote_value
  partner_sum
  partner_full_name
  partner_total_quotes
  total_quotes
  partner_sum
  percentage
  sum_percentage
]

# Run the debugger with ALL placeholders (no .uniq) - test each one individually
debugger = Docx::Debugger.analyze('tests/CS-TEMPLATE-mus.docx') do |config|
  config.placeholder_type = :mustache
  config.set_placeholders(placeholders)
end

debugger.debug!
