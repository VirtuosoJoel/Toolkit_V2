require 'date'

class FiscalYear
  
  FiscalYearStart = Date.new( 0, 9, 26 )
  
  attr_reader :months, :weeks, :days, :year
  
  def initialize( starting_year )
  
    @year = starting_year
    @days = (Date.new(starting_year, 9, 26)..Date.new(starting_year+1, 9, 25)).to_a
    @weeks = days.each_slice(7).to_a
    
    temp = weeks.dup
    @months = ([4,4,5]*3 + [4,4,6]).map { |num| temp.shift(num) }
    
  end

  # Find which fiscal month a date belongs to
  def month( date )
    (0...months.size).find { |i| months[i].any? {|week| week.include? date } } + 1
  end
  
  # Find which fiscal week a date belongs to
  def week( date )
    (0...weeks.size).find { |i| weeks[i].include? date } + 1
  end
  
  # String representation of the fiscal year
  def to_s
    "FiscalYear #{ year }: #{ days.first.strftime('%d/%m/%Y') } to #{ days.last.strftime('%d/%m/%Y') }"
  end
  
  # Return the current fiscal year for the given date or today
  def self.current( target_date = Date.today )
    FiscalYear.new( target_date.year - ( Date.new(0, target_date.month, target_date.day ) < FiscalYearStart ? 1 : 0 ) )
  end
  
end # FiscalYear
