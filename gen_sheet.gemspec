Gem::Specification.new do |s|
  s.name        = 'gen_sheet'
  s.version     = '0.0.4'
  s.summary     = "A spreadsheet generator (ODS, XLS) for Roo."
  s.description = "Takes a Roo spreadsheet object and renders it in XLS or ODS. Also allows you to use a template sheet and fill it with data from Roo, then output a rendered sheet."
  s.authors     = ["Rei Kagetsuki"]
  s.email       = 'zero@genshin.org'
  s.files        = `git ls-files`.split("\n")
  s.homepage    = 'https://github.com/Genshin/GenSheet'

  s.add_dependency 'roo'
  s.add_dependency 'writeexcel'
  s.add_dependency 'rodf'
end
