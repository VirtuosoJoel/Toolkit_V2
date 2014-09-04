require 'win32ole'
require_relative 'MailAddressFormatter'

#
# Note: The mail methods "Display" and "Send" are case-sensitive due to Ruby methods with the same name.
#

class Outlook
  extend MailAddressFormatter
  
  attr_accessor :application, :mail
 
  def initialize
    self.application = WIN32OLE.new('Outlook.Application')
    defined?( OlMailItem ) or WIN32OLE.const_load(application, Outlook)
  end
  
  def create_mail
    self.mail = application.CreateItem(OlMailItem)
  end

  def send_email( files: [], subject: 'Automated Email', to: MyName, body: 'Automated Email', from: EmailFrom, cc: nil, bcc: nil )
    
    create_mail
    
    mail.to = Outlook.format_addresses( to )
    mail.cc = Outlook.format_addresses(cc)
    mail.bcc = Outlook.format_addresses(bcc)

    files.each { |file| mail.attachments.add file }
    
    mail.subject = subject
    
    mail.htmlbody = body
    
    mail.Send
  end
  
end
