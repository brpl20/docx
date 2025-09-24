#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'docx'
require 'json'

class Formatter
  ENTITY_PREFIXES = {
    person: "residente e domiciliado",
    company: "com sede a"
  }.freeze

  ADDRESS_FORMAT = {
    city_state_separator: ' - ',
    zip_prefix: 'CEP: '
  }.freeze

  DOCUMENT_PREFIXES = {
    cpf: {
      male: "inscrito no CPF sob o nº",
      female: "inscrita no CPF sob o nº",
    },
    cnpj: {
      company: "inscrita no CNPJ sob o nº"
    }
  }.freeze

  def initialize(data, entity_type = :person)
    @data = data
    @entity_type = entity_type
    @gender = (val(:gender) || :male).to_sym
  end

  def self.full_name(data, entity_type = :person)
    new(data, entity_type).full_name
  end

  def self.address(data, entity_type = :person)
    new(data, entity_type).address
  end

  def self.cpf(data, entity_type = :person)
    new(data, entity_type).cpf
  end

  def self.cnpj(data, entity_type = :company)
    new(data, entity_type).cnpj
  end

  def full_name
    clean_join(val(:name), val(:last_name)).upcase
  end

  def cpf
    return unless val(:cpf)
    gender_key = @entity_type == :company ? :company : @gender
    prefix = DOCUMENT_PREFIXES[:cpf][gender_key]
    "#{prefix} #{val(:cpf)}"
  end

  def cnpj
    return unless val(:cnpj)
    prefix = DOCUMENT_PREFIXES[:cnpj][:company]
    "#{prefix} #{val(:cnpj)}"
  end

  def rg
    return unless val(:rg)
    gender_prefix = @gender == :female ? "portadora" : "portador"
    "#{gender_prefix} da cédula de identidade RG nº #{val(:rg)}"
  end

  def oab
  end


  def email
    val(:email)
  end

  def phone
    val(:phone)
  end

  def address
    prefix = ENTITY_PREFIXES[@entity_type]
    street_number = [val(:street), val(:number)].compact.join(', ')
    city_state = "#{val(:city)}#{ADDRESS_FORMAT[:city_state_separator]}#{val(:state)}"
    zip_with_prefix = val(:zip) ? "#{ADDRESS_FORMAT[:zip_prefix]}#{val(:zip)}" : nil
    parts = [street_number, city_state, zip_with_prefix].compact.reject(&:empty?)

    "#{prefix} #{parts.join(', ')}"
  end

  def profession
  end

  def marital_status
  end

  def nationality
  end

  def nationality
  end

  def qualification
    full_name
    nationality
    marital_status
    profession
    cpf
    rg
    oab if exist
    address
  end



private

  def val(key)
    @data[key] || @data[key.to_s]
  end

  def clean_join(*parts)
    parts
      .compact
      .map(&:to_s)
      .join(' ')
      .gsub(/\s+/, ' ') # Rails => Replace for: Squish
      .strip
  end
end
