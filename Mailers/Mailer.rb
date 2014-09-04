require 'mail'
require 'nokogiri'
require_relative 'MailAddressFormatter'


class Mailer
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

  def send_email( files: [], subject: 'Automated Email', to: MyName, body: 'Automated Email', from: EmailFrom, cc: nil, bcc: nil )
    
    # Grab the parameters as a Hash as well as the above variables
    params = Hash[method(__method__).parameters.map { |_,v| [v, eval(v.to_s)] }]
    
    # Create a new email
    m = Mail.new
    
    # Email addresses
    [ :from, :to, :cc, :bcc ].each do |address_type|
      m[ address_type ] = Mailer.format_addresses( params[ address_type ] )
    end
    
    # Attachments
    files.each { |f| m.add_file f }
    
    # Subject
    m[ :subject ] = subject
    
    # HTML Part
    m.text_part = Mail::Part.new do
        content_type 'text/html; charset=UTF-8'
        body params[ :body ]
    end

    # Text Part
    m.html_part = Mail::Part.new do
      content_type 'text/plain; charset=UTF-8'
      body Mailer.html_to_text( params[:body] )
    end
  
    # Send the email
    m.deliver!
    
  end # send_email
  
  def self.html_to_text( html )
  
    doc = Nokogiri::HTML(html)
    doc.css('script, link').each &:remove
    doc.css('body').text.squeeze(" \n")
  
  end
  
end
