require 'date'

class FiscalYear
  
  ReferencePoint = Date.new( 1900, 10, 25 )
  
  attr_reader :months, :weeks, :days, :year, :start_date
  
  def initialize( starting_year )
  
    @year = starting_year
  
    @start_date = ReferencePoint
    (ReferencePoint.year...year).each { |y| @start_date += year_days( y ) }
    
    @days = (start_date...(start_date+year_days)).to_a
    @weeks = days.each_slice(7).to_a
    
    temp = weeks.dup
    @months = (week_pattern).map { |num| temp.shift(num) }
    
  end

  def year_days( target_year=year )
    year_weeks( target_year ) * 7
  end
  
  def year_weeks( target_year=year )
    target_year % 7 == 5 ? 53 : 52
  end
  
  def leap?
    Date.leap?(year+1)
  end
  
  # String representation of the Fiscal Year in grid format
  def display
    a = month_week_pattern
    weeks.map.with_index { |week,idx| "Month: #{ '%02d' % (a[idx]) }\tWeek: #{ '%02d' % (idx+1) }\t#{ fd(week.first) }\t#{ fd(week.last) }" }
  end
  
  # Standardise the date format without repeating code
  def format_date( date = Date.today )
    date.strftime('%a %d/%m/%Y')
  end
  alias fd format_date
  
  # Find which fiscal month a date belongs to
  def month( date = Date.today )
    (0...months.size).find { |i| months[i].any? {|week| week.include? date } } + 1
  end
  
  # Find which fiscal week a date belongs to
  def week( date = Date.today )
    (0...weeks.size).find { |i| weeks[i].include? date } + 1
  end
  
  # Fiscal Month 4,4,5 week pattern
  def week_pattern
    year_weeks == 53 ? [5,4,5]+[4,4,5]*3 : [4,4,5]*4
  end
  
  # The Month each week belongs to in Array form
  def month_week_pattern
    week_pattern.map.with_index { |val,idx| Array.new(val,idx+1) }.flatten
  end
  
  # String representation of the Fiscal Year
  def to_s
    "FiscalYear #{ year }: #{ fd(days.first) } to #{ fd(days.last) }"
  end
  
  # Return the current fiscal year for the given date or today
  def self.current( target_date = Date.today )
    FiscalYear.new( target_date.year - ( Date.new(0, target_date.month, target_date.day ) < FiscalYearStart ? 1 : 0 ) )
  end
  
end # FiscalYear
