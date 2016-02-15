require 'spreadsheet'
require 'byebug'
require 'json'

class SnacCodeMatcher
  AUTHORITIES_FILE= "authorities.json"
  FILENAME = "test.xls"

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
    return "NOT FOUND" if authorities[trimmed_name(name)].nil?
    authorities[trimmed_name(name)]["ons"]
  end

  def trimmed_name(name)
    name.gsub(' Borough', '').gsub(' City', '').gsub(' County', '').downcase.gsub(' ', '-')
  end

  def authorities
    @_authorities ||= JSON.parse(authorities_file)
  end

  def authorities_file
    @_file ||= File.read(AUTHORITIES_FILE)
  end

  def names
    [
      "Barking and Dagenham",
      "Barnet",
      "Barnsley",
      "Bath and North East Somerset",
      "Central Bedfordshire",
      "Bedford Borough",
      "Bexley",
      "Birmingham",
      "Blackburn with Darwen",
      "Blackpool",
      "Bolton",
      "Bournemouth",
      "Bracknell Forest",
      "Bradford Metropolitan District",
      "Brent",
      "Brighton and Hove",
      "Bristol City",
      "Bromley",
      "Buckinghamshire County",
      "Bury",
      "Calderdale",
      "Cambridgeshire",
      "Camden",
      "Cheshire East",
      "Cheshire West and Chester",
      "City of London",
      "Cornwall and Isles of Scilly",
      "Coventry",
      "Croydon",
      "Cumbria",
      "Darlington",
      "Derby City",
      "Derbyshire",
      "Devon",
      "Doncaster",
      "Dorset",
      "Dudley",
      "Durham",
      "Ealing",
      "East Riding of Yorkshire",
      "East Sussex",
      "Enfield",
      "Essex",
      "Gateshead",
      "Gloucestershire",
      "Greenwich",
      "Hackney and City",
      "Halton",
      "Hammersmith and Fulham",
      "Hampshire",
      "Haringey",
      "Harrow",
      "Hartlepool",
      "Havering",
      "Herefordshire",
      "Hertfordshire",
      "Hillingdon",
      "Hounslow",
      "Isle Of Wight",
      "Isles of Scilly - see Cornwall & Isles of Scilly",
      "Islington",
      "Kensington and Chelsea",
      "Kent",
      "Kingston Upon Hull",
      "Kingston Upon Thames",
      "Kirklees",
      "Knowsley",
      "Lambeth",
      "Lancashire",
      "Leeds City",
      "Leicester City",
      "Leicestershire and Rutland",
      "Lewisham",
      "Lincolnshire",
      "Liverpool",
      "Luton",
      "Manchester",
      "Medway Towns",
      "Merton",
      "Middlesbrough",
      "Milton Keynes",
      "Newcastle",
      "Newham",
      "Norfolk",
      "North East Lincolnshire",
      "North Lincolnshire",
      "North Somerset",
      "North Tyneside",
      "North Yorkshire",
      "Northamptonshire",
      "Northumberland",
      "Nottingham City",
      "Nottinghamshire",
      "Oldham",
      "Oxfordshire",
      "Peterborough",
      "Plymouth",
      "Poole",
      "Portsmouth",
      "Reading",
      "Redbridge",
      "Redcar and Cleveland",
      "Richmond Upon Thames",
      "Rochdale",
      "Rotherham",
      "Rutland",
      "Salford",
      "Sandwell",
      "Sefton",
      "Sheffield",
      "Shropshire",
      "Slough",
      "Solihull",
      "Somerset",
      "South Gloucestershire",
      "South Tyneside",
      "Southampton",
      "Southend  ",
      "Southwark",
      "St Helens",
      "Staffordshire",
      "Stockport",
      "Stockton on Tees",
      "Stoke On Trent",
      "Suffolk",
      "Sunderland",
      "Surrey",
      "Sutton",
      "Swindon",
      "Tameside",
      "Telford and Wrekin",
      "Thurrock",
      "Torbay",
      "Tower Hamlets",
      "Trafford",
      "Wakefield",
      "Walsall",
      "Waltham Forest",
      "Wandsworth",
      "Warrington",
      "Warwickshire",
      "West Berkshire",
      "West Sussex",
      "Westminster",
      "Wigan",
      "Wiltshire",
      "Windsor and Maidenhead",
      "Wirral",
      "Wokingham",
      "Wolverhampton",
      "Worcestershire",
      "York City",
    ]
  end
end

SnacCodeMatcher.new.write_to_xls
