require 'nokogiri'
require 'win32ole'
require_relative 'MailAddressFormatter'

class CDO
  extend MailAddressFormatter
  include Passwords
 
  def initialize
  
    

    Mail.defaults do
      delivery_method :smtp,
      { 
        :address              => EmailServer,
        :port                 => 25,
        :domain               => EmailDomain
      }
      
    end

  end

  def create_mail

    m = WIN32OLE.new('CDO.Message') 
    
    [ 
      [ 'http://schemas.microsoft.com/cdo/configuration/sendusing', 2 ],
      [ 'http://schemas.microsoft.com/cdo/configuration/smtpserver', EmailServer ],
      [ 'http://schemas.microsoft.com/cdo/configuration/smtpserverport', 25 ]
    ].each do |item_name, item_value|
      m.Configuration.Fields.Item( item_name ).Value = item_value
    end
    m.Configuration.Fields.Update
    
    m
  
  end
  
  def send_email( files: [], subject: 'Automated Email', to: MyName, body: 'Automated Email', from: EmailFrom, cc: nil, bcc: nil )
    
    # Grab the parameters as a Hash as well as the above variables
    params = Hash[method(__method__).parameters.map { |_,v| [v, eval(v.to_s)] }]
    
    # Create a new email
    m = create_mail
    
    # Email addresses
    [ :from, :to, :cc, :bcc ].each do |address_type|
      m.send( address_type.to_s + '=', Mailer.format_addresses( params[ address_type ] ) )
    end
    
    p m.from, m.to, m.cc, m.bcc
    
    # Attachments
    files.each { |f| m.AddAttachment f }
    
    # Subject
    m.Subject = subject
    
    # HTML Part
    m.HTMLBody = body

    # Text Part
    m.TextBody = Mailer.html_to_text( body )
  
    # Send the email
    m.Send
    
  end # send_email
  
  def self.html_to_text( html )
  
    doc = Nokogiri::HTML(html)
    doc.css('script, link').each &:remove
    doc.css('body').text.squeeze(" \n")
  
  end
  
end
