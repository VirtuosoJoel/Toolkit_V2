require 'win32ole'
require 'rubyexcel'
require_relative 'RegistryTools'
require_relative 'DateTools'

class RoleCall
  attr_accessor :cn, :database, :name, :result, :rs, :sql, :placeholders, :excel, :wb
  alias connection cn
  alias recordset rs

  SQL_STRINGS = {
    Select: "Select TOP 1 SQL FROM Dashboard.dbo.VBA_Report_Queries WHERE name = 'REPORT_NAME_PLACEHOLDER'".freeze,
    Update: "UPDATE Dashboard.dbo.VBA_Report_Queries SET SQL = 'SQL_PLACEHOLDER', Description = 'DESCRIPTION_PLACEHOLDER' WHERE Name = 'NAME_PLACEHOLDER'".freeze,
    Insert: "INSERT INTO Dashboard.dbo.VBA_Report_Queries(name,sql,description)VALUES('NAME_PLACEHOLDER','SQL_PLACEHOLDER','DESCRIPTION_PLACEHOLDER')".freeze
  }.freeze
  
  CONNECTION_STRINGS = {
    RoleCall: "Provider=SQLOLEDB;User ID=#{ Passwords::SQLRoleCall::UserName };password=#{ Passwords::SQLRoleCall::Password };Initial Catalog=ROLECALL;Data Source=#{ Passwords::SQLRoleCall::DSN };".freeze,
    Mitel: "Provider=SQLOLEDB;User ID=#{ Passwords::SQLMitel::UserName };Password=#{ Passwords::SQLMitel::Password };Initial Catalog=CCMData;Data Source=#{ Passwords::SQLMitel::DSN };".freeze,
    AlarmMaster: "Dsn=#{ Passwords::SQLAlarmMaster::DSN };uid=#{ Passwords::SQLAlarmMaster::UserName }".freeze
  }.freeze

  def initialize
    self.database = :RoleCall
    self.cn = WIN32OLE.new('ADODB.Connection')
    self.rs = WIN32OLE.new('ADODB.Recordset')
    self.placeholders = {}
  end
  
  def clean_name
    name.gsub(/(\AGet_)/i,'').gsub(/_/,' ')[0..30]
  end
  
  def clear
    self.name = nil
    self.result = nil
  end
  
  def connect( connection_name = database )
    disconnect
    cn.Open( CONNECTION_STRINGS[ connection_name ] ); self
  end
  alias open connect

  def disconnect
    [ cn, rs ].each { |obj| obj.close unless obj.state.zero? }; self
  end
  alias close disconnect
  
  def display
    result ? result.to_excel : raise( NoMethodError, 'No result to display.' )
  end
  
  def execute( command )
    connect
    cn.execute command
  end
  
  def get_sql
    fail NoMethodError, 'No report name specified' unless name
    connect :RoleCall
    self.sql = query( SQL_STRINGS[ :Select ].gsub( 'REPORT_NAME_PLACEHOLDER', name ) ).A2
    sql
  end
  
  # Overwrite the name setter so it wipes the SQL code when set
  def name=( str )
    self.sql = nil
    @name = str
  end
  
  # Run a specific query and return a RubyExcel Sheet
  def query( sql_string )
    
    # Connect to the Server
    connect
    
    # Bring the Data into a RecordSet
    rs.open( sql_string, cn )
    
    # Get the Headers
    fields = rs.Fields.each.map &:name
    
    # Turn the Data into a 2D Array
    data = rs.GetRows.transpose rescue [[]]
    
    # Load the headers & data into a RubyExcel Sheet
    RubyExcel::Workbook.new.load( [fields] + data )

  end
  
  # Run a query based on the report name
  def run
    connect if cn.state.zero?
    sql || get_sql
    replace_placeholders
    
    # If we already hold a Sheet
    if result 
      
      # Take the Workbook which holds the current Sheet
      rwb = result.parent
      
      # Run the Report and bring back a new Sheet
      self.result = query( sql )
      
      # Add the new Sheet onto the existing Workbook
      rwb << result
      
    else
    
      # Create a new Sheet
      self.result = query( sql )
      
    end
    
    # Use the SQL report naming convention to guess at a sheet name.
    # Max Sheet name length in Excel is 31 characters.
    result.name = clean_name if name
    
    # Set the result to be invisible by default
    result.parent.standalone = true
    
    # Return the Sheet
    result
    
  end
  
  def run_report( report_name, sheet_name )
    self.name = report_name
    run
    result.name = sheet_name
  end
  
  def replace_placeholders
    placeholders.each { |find_str, replace_with_str| self.sql = sql.gsub( find_str, RoleCall.escape_quotes( replace_with_str ) ) }
  end

  # Save the current results as an Excel Workbook and store handles to both the Workbook and Excel
  def save( filename = clean_name )
    self.wb = result.parent.save_excel( filename, true )
    self.excel = wb.application
    wb
  end
  alias saveas save
  
  # Save the current results and then close the Workbook, and Excel (if excel is not visible)
  # Return the full filename
  def save_and_close( filename = clean_name )
    result.parent.export filename
  end
  alias export save_and_close
  
  def self.escape_quotes( param )
    param.to_s.gsub(/'/, "''")
  end
  
end

=begin
# Example usage

rc = RoleCall.new
rc.placeholders = {
  'FISCAL_YEAR_START_PLACEHOLDER' => 2013,
  'WEEK_PLACEHOLDER' => 49
}
rc.name = 'Get_Week_Dates'
rc.run
rc.export

=end
