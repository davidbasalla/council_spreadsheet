require 'spreadsheet'
require 'byebug'
require 'json'
require 'csv'

class SnacCodeMatcher
  AUTHORITIES_FILE= "authorities.json"
  PLACES_FILE = "register_offices.csv"
  FILENAME = "test_2.xls"

  def initialize
  end

  def write_to_xls
    worksheet
    insert_snac_codes
    book.write(FILENAME)
  end

  def insert_snac_codes
    names.each_with_index do |n, index|
      book.worksheet(0).insert_row(index, [n, match_snac_code(n)])
    end
  end

  def book
    @book ||= Spreadsheet::Workbook.new
  end

  def worksheet
    @worksheet ||= book.create_worksheet :name => 'Sheet Name'
  end

  def match_snac_code(name)
    return "NOT FOUND: #{name}" if authority_match(name).nil?
    authorities[trimmed_name(name)]["ons"]
  end

  def authority_match(name)
    authorities[trimmed_name(name)]
  end

  def trimmed_name(name)
    name
    .gsub('London Borough of ','')
    .gsub(' - Register Office', '')
    .gsub(' - Regiser Office', '')
    .gsub(' Register Office', '')
    .gsub(' - Registration Office', '')
    .gsub(' Registration Office', '')
    .gsub(' Council', '')
    .gsub(' Borough', '')
    .gsub(' Metropolitan', '')
    .gsub(' District', '')
    .gsub(' City', '')
    .gsub(' County', '')
    .downcase.gsub(' ', '-')
  end

  def authorities
    @_authorities ||= JSON.parse(authorities_file)
  end

  def authorities_file
    @_authorities_file ||= File.read(AUTHORITIES_FILE)
  end

  def places_file
    @_places_file ||= File.read(PLACES_FILE)
  end

  def names
    @_names ||= CSV.new(places_file).to_a.map { |e| e[0] }
  end
end

SnacCodeMatcher.new.write_to_xls
