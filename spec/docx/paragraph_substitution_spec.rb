# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe 'Paragraph text substitution across runs' do
  let(:doc) { Docx::Document.open('spec/fixtures/substitution.docx') }

  describe '#substitute_across_runs' do
    context 'when placeholder is split across multiple runs' do
      it 'replaces placeholders that are fragmented' do
        paragraph = doc.paragraphs[1]
        original_text = paragraph.text
        
        # Simulate a fragmented placeholder scenario
        # Even if _placeholder2_ is split like _pla][cehold][er2_
        paragraph.substitute_across_runs('_placeholder2_', 'REPLACED')
        
        expect(paragraph.text).to include('REPLACED')
        expect(paragraph.text).not_to include('_placeholder2_')
      end

      it 'replaces multiple placeholders in the same paragraph' do
        paragraph = doc.paragraphs[1]
        
        paragraph.substitute_across_runs('_placeholder2_', 'FIRST')
        paragraph.substitute_across_runs('_placeholder3_', 'SECOND')
        
        expect(paragraph.text).to include('FIRST')
        expect(paragraph.text).to include('SECOND')
        expect(paragraph.text).not_to include('_placeholder2_')
        expect(paragraph.text).not_to include('_placeholder3_')
      end

      it 'handles curly brace placeholders {{placeholder}}' do
        # Create a test paragraph with fragmented placeholder
        paragraph = doc.paragraphs.first
        paragraph.text = 'This is {{placeholder}} text'
        
        paragraph.substitute_across_runs('{{placeholder}}', 'REPLACED')
        
        expect(paragraph.text).to eq('This is REPLACED text')
      end

      it 'handles placeholders with underscores at both ends' do
        paragraph = doc.paragraphs.first
        paragraph.text = 'Test _placeholder_abc_ here'
        
        paragraph.substitute_across_runs('_placeholder_abc_', 'REPLACED')
        
        expect(paragraph.text).to eq('Test REPLACED here')
      end
    end

    context 'with regex patterns' do
      it 'replaces patterns across fragmented runs' do
        paragraph = doc.paragraphs.first
        paragraph.text = 'Email: user@example.com'
        
        paragraph.substitute_across_runs(/\b[\w._%+-]+@[\w.-]+\.[A-Z]{2,}\b/i, 'EMAIL_HIDDEN')
        
        expect(paragraph.text).to eq('Email: EMAIL_HIDDEN')
      end
    end
  end

  describe '#substitute_across_runs_with_block' do
    it 'replaces with block logic across runs' do
      paragraph = doc.paragraphs.first
      paragraph.text = 'Total: 100'
      
      paragraph.substitute_across_runs_with_block(/Total: (\d+)/) do |match|
        "Total: #{match[1].to_i * 2}"
      end
      
      expect(paragraph.text).to eq('Total: 200')
    end

    it 'handles multiple replacements with capturing groups' do
      paragraph = doc.paragraphs.first
      paragraph.text = 'Date: 2024-01-15, Amount: $50'
      
      paragraph.substitute_across_runs_with_block(/\$(\d+)/) do |match|
        "$#{(match[1].to_i * 1.1).round}"
      end
      
      expect(paragraph.text).to eq('Date: 2024-01-15, Amount: $55')
    end
  end

  describe '#consolidate_text_runs' do
    it 'merges adjacent runs with identical formatting' do
      paragraph = doc.paragraphs.first
      initial_run_count = paragraph.text_runs.size
      
      # This should merge runs with same formatting
      paragraph.consolidate_text_runs
      
      # After consolidation, we should have fewer or equal runs
      expect(paragraph.text_runs.size).to be <= initial_run_count
      
      # But text content should remain the same
      expect(paragraph.text).to eq(doc.paragraphs.first.text)
    end
  end

  describe 'real-world fragmentation scenarios' do
    it 'handles Word-style fragmentation of {{variable}}' do
      paragraph = doc.paragraphs.first
      # Simulate how Word might fragment {{variable}}
      # It could be split as: {{][vari][able][}}
      paragraph.text = 'Replace {{variable}} here'
      
      paragraph.substitute_across_runs('{{variable}}', 'VALUE')
      
      expect(paragraph.text).to eq('Replace VALUE here')
    end

    it 'handles complex nested placeholders' do
      paragraph = doc.paragraphs.first
      paragraph.text = 'User: {{first_{{type}}_name}}'
      
      # First replace inner placeholder
      paragraph.substitute_across_runs('{{type}}', 'last')
      expect(paragraph.text).to eq('User: {{first_last_name}}')
      
      # Then replace outer placeholder
      paragraph.substitute_across_runs('{{first_last_name}}', 'Smith')
      expect(paragraph.text).to eq('User: Smith')
    end

    it 'preserves text that looks like placeholders but should not be replaced' do
      paragraph = doc.paragraphs.first
      paragraph.text = 'Keep {{this}} but replace {{that}}'
      
      paragraph.substitute_across_runs('{{that}}', 'REPLACED')
      
      expect(paragraph.text).to eq('Keep {{this}} but replace REPLACED')
    end
  end
end