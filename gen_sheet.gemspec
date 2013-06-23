Gem::Specification.new do |s|
  s.name        = 'gen_sheet'
  s.version     = '0.0.5'
  s.summary     = "A spreadsheet generator (ODS, XLS) and parsing tool. Uses Roo as a backend and attempts to implement newer features on top of it."
  s.description = "A spreadsheet generator (ODS, XLS) and parsing tool. Uses Roo as a backend and attempts to implement newer features on top of it."
  s.authors     = ["Rei Kagetsuki", "Rika Yoshida"]
  s.license     = ["LGPL"]
  s.email       = 'zero@genshin.org'
  s.files        = `git ls-files`.split("\n")
  s.homepage    = 'https://github.com/Genshin/GenSheet'

  s.add_dependency 'roo'
  s.add_dependency 'writeexcel'
  s.add_dependency 'rodf'
end
