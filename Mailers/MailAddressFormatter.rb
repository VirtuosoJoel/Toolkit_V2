require_relative '../Passwords'

module MailAddressFormatter
  include Passwords

  def format_addresses( address_list )
  
    return '' if address_list.to_s.empty?
    
    [address_list].flatten.map do |email_address|
      Mailer.format_address email_address
    end.join(';')
    
  end
  
  def format_address( email_address )
    email_address.include?('@') ? email_address : email_address.gsub(/\s+/,'.') + EmailSuffix
  end
  
end
