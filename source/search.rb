#!/usr/bin/env ruby

# TODO Add support to quicklook contacts using Outlooks Vcard data, pseudo code:
# fork do
#   # Get 'vcard data' and save to a temporary file, use 'Dir.mktmpdir <pathprefix>'
#   # then we can set the quicklookurl to the temp file,
#     see https://www.alfredapp.com/help/workflows/inputs/script-filter/json/
# end

require 'json'

ICONS = {
  manager: "/Applications/Microsoft Outlook.app/Contents/Resources/ol_Reminder_Icon_people_16px@2x.png",
  direct_report: "/Applications/Microsoft Outlook.app/Contents/Resources/ol_user@2x.png",
  contact: "/Applications/Microsoft Outlook.app/Contents/Resources/ol_Bar_icons_People_ToolbarMode_hover@2x.png"
}

class String
  def titlecase
    gsub(/\s+/, ' ')
      .split(/([[:alpha:]]+)/)
      .map(&:capitalize)
      .join
  end

  # Applescript returns 'missing value', so treat those as a nil
  def nil?
    self == 'missing value'
  end
end

class Contact
  attr_reader :name, :email, :role

  def initialize(contact_name, job_title, email = nil, role:)
    @email, @job_title, @role = email, job_title, role

    self.name = contact_name
  end

  def name=(name)
    @name ||= if ENV['NAME_FILTER']
                regexp = Regexp.new(ENV['NAME_FILTER'])
                name.gsub(regexp, '')
                    .strip
              else
                name
              end
  end

  def display
    @display ||= if @name.nil?
                   @email
                 elsif @job_title.nil?
                   @name
                 else
                   "#{@name} • #{@job_title.titlecase}"
                 end

  end

  def first_name
    @first_name ||= @name.split.first
  end

  def last_name
    @name.split[1..-1].join ' '
  end

  def role_name
    @role.to_s.sub('_', ' ').titlecase
  end

  # Our organization does not use the 'IM addresses' field in Exchange.
  # Instead a variation of the email address is used.
  #
  # File a bug report if you'd like a different im address handling.
  def im_address
    email.sub /[^@]+/, (first_name[0] + last_name).downcase
  end
end

def contacts_builder(results, role)
  regex = role.to_s << ':"(?<name>[^"]+)", title:"(?<title>[^"]+)"'
  regex << ', address:"(?<email>[^",]+)", ' if role == :contact

  regexp = Regexp.new regex
  results.scan(regexp)
         .collect { |matches| Contact.new *matches, role: role }
end

def empty_result
  {
    items: [{
             title: "No contacts found",
             icon: {
               path: "/Applications/Microsoft Outlook.app/Contents/Resources/ol_Bar_icons_People_ToolbarMode_press@2x.png"
             },
             valid: false
           }]
  }
end

def alfred_entry(contact)
  if contact.role == :contact
    {
      uid: contact.email,
      title: contact.display,
      subtitle: "✉️ #{contact.email}",
      arg: "mailto:#{contact.name} <#{contact.email}>",
      icon: {
        path: ICONS[contact.role]
      },
      autocomplete: contact.name,
      text: {
        copy: contact.email,
        largetype: contact.email
      },
      match: contact.name,
      mods: {
        shift: {
          arg: contact.email,
          subtitle: "Open #{contact.first_name}’s contact card"
        },
        alt: {
          arg: contact.name,
          subtitle: "Search Spotlight for ‘#{contact.name}’"
        },
        cmd: {
          arg: contact.im_address,
          subtitle: "💬 #{contact.im_address}"
        }
      }
    }
  else
    {
      title: contact.role_name,
      subtitle: contact.display,
      arg: contact.name,
      icon: {
        path: ICONS[contact.role]
      },
      autocomplete: contact.name,
      valid: false,
      mods: {
        alt: {
          valid: true,
          arg: contact.name,
          subtitle: "Search for '#{contact.name}' in Spotlight"
        }
      }
    }
  end
end

def build_alfred_results(contacts, org_hierarchy)
  {
    items: contacts.collect { |c| alfred_entry(c) } + org_hierarchy.collect { |c| alfred_entry(c) }
  }
end

def search_exchange(query)
  results = `/usr/bin/osascript "./search.applescript" #{query}`
  contacts_builder(results, :contact)
end

def search_exchange_hierarchy(contact)
  contact_detail = `/usr/bin/osascript "./hierarchy.applescript" #{contact}`
  [:manager, :direct_report].collect { |role| contacts_builder(contact_detail, role) }
                            .flatten
end

def main
  contacts = search_exchange($*.join(' '))
  org_hierarchy = contacts.one? ? search_exchange_hierarchy(contacts.first.email) : []

  alfred_results = contacts.any? ? build_alfred_results(contacts, org_hierarchy) : empty_result

  puts JSON.fast_generate(alfred_results)
end

main
