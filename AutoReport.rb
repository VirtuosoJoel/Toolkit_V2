
# Class to simplify automated reporting
class AutoReport

  # Bring in the Report code
  require_relative 'RoleCall'
  
  # Bring in the Outlook Email code
  require_relative 'Mailers/Outlook'
  
  # Bring in the SMTP Email code
  require_relative 'Mailers/Mailer'
  
  include DateTools
  include RegistryTools
  
  # Provide access to Class contents
  attr_accessor :rc, :mailer, :name, :to, :files, :body, :checkname, :subject, :automation, :test, :mailer, :cancel, :from, :bcc, :cc

  # Set up the Class
  def initialize( to = Passwords::MyName )
  
    # The SMTP server will not relay to external email addresses, so we use Outlook for external emailing
    # The user can specify a mailer using the attr_accessor if required
    if [to].flatten.any? { |addr| addr.include?( '@' ) && addr !~ /#{ Regexp.quote(Passwords::EmailSuffix) }/ }
      self.mailer = Outlook.new
    else
      self.mailer = Mailer.new
    end
  
    # Reporting tool
    self.rc = RoleCall.new
    
    # Report name
    self.name = File.basename( $0, '.*' )
    
    # File to check for failed reports
    self.checkname = "#{ RubyExcel.documents_path }\\#@name.txt"
    
    # Email content
    self.to = to
    self.files = []
    self.body = 'Automated Email'
    self.subject = nil
    self.from = EmailFrom
    self.cc = nil
    self.bcc = nil
    
    # Check automation mode
    case ARGV[0]
    when /check/i
      puts 'Check mode active. Setting Automate mode: true'
      self.automation = true
      if ( File.read( checkname ) == today.to_s rescue false )
        puts 'Report already completed today'
        exit
      end
    when /automate/i
      puts 'Automated'
      self.automation = true
    when /test/i
      self.automation = true
      self.test = true
      self.to = Passwords::MyName
      self.cc = nil
      self.bcc = nil
    else
      puts 'Not Automated'
      self.automation = false
    end
    
    # Cancel option
    self.cancel = false
    
  end # initialize

  # Wrapper for the report code, automatically handles errors and automated emailing
  def run
    tried = false
  
    # Error handler
    begin
    
      # This is where the real work gets done
      yield rc
      
    # Catch errors
    rescue => err
    
      # Display error in console
      puts ( error_details = "#{ err.message }\n\n#{ err.backtrace.join($/) }" )
    
      if tried
      
        # Email failure message if automated
        logerror( "#{ Time.now }\n#{ error_details }" ) if automation && !test
        exit
        
      else
      
        tried = true
        
        # Only retry if automated
        if automation
          retry
        else
          # Only hold the console window open if not automated
          gets
          exit
        end
        
      end # tried
      
    end # errorhandler
    
    if automation && !cancel
      
      begin
      
        # Send email and log success if automated and successful
        sendmail
        File.write( checkname, today )
      
      # Catch failures like invalid email addresses
      rescue => err
        
        # Display error in console
        puts ( error_details = "#{ err.message }\n\n#{ err.backtrace.join($/) }" )
        
        # Email failure message
        logerror( "#{ Time.now }\n#{ error_details }" ) if automation && !test
        raise err
        
      end # errorhandler
      
    end # if automation
    
  end # run
  
  def sendmail
    options = {
      files: files,
      subject: ( subject || name ),
      to: ( test ? Passwords::MyName : to ),
      body: body,
      from: from,
      bcc: ( test ? nil : bcc ),
      cc: ( test ? nil : cc )
    }
    mailer.send_email options
  end
  
  def logerror( error )
    options = {
      files: [],
      subject: "Failure - #{ name } #{ date_str }",
      to: Passwords::ErrorAlert,
      body: error
    }
    mailer.send_email options
  end
  
end
