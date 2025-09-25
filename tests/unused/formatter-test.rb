require_relative 'formatter'

p1 = { 'name' => 'Ada', :last_name => 'Lovelace', gender: 'female', cpf: '123.456.789-00', rg: '12.345.678-9' }
p2 = { name: 'Alan', last_name: 'Turing', gender: 'male', cpf: '987.654.321-00' }
p3 = { street: "Rua Alexandre de Gusmão", number: 712, city: "Cascavel", state: "PR", zip: "85819-530" }
company = { name: 'Tech Corp', cnpj: '12.345.678/0001-90', street: "Av. Brasil", number: 1000, city: "São Paulo", state: "SP" }

puts "=== Full Names ==="
puts Formatter.full_name(p1)  # => "ADA LOVELACE"
puts Formatter.full_name(p2)  # => "ALAN TURING"

puts "\n=== CPF/CNPJ ==="
puts Formatter.cpf(p1)  # Female: "inscrita no CPF sob o nº 123.456.789-00"
puts Formatter.cpf(p2)  # Male: "inscrito no CPF sob o nº 987.654.321-00"
puts Formatter.cnpj(company, :company)  # "inscrita no CNPJ sob o nº 12.345.678/0001-90"

puts "\n=== RG ==="
puts Formatter.new(p1).rg  # Female: "portadora da cédula de identidade RG nº 12.345.678-9"
puts Formatter.new(p2).rg if p2[:rg]  # No RG, won't print

puts "\n=== Addresses ==="
puts "Person address:"
puts Formatter.address(p3, :person)
puts "\nCompany address:"
puts Formatter.address(company, :company)
