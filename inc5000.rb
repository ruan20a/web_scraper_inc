require "rubygems"
require "nokogiri"
require "mechanize"
require "open-uri"
require "capybara"
require "pp"
require "pry"
require 'pry-byebug'
require "timeout"
require "writeexcel"

base_url = "http://www.inc.com/inc5000/list/2014"

def make_angular_ready(session)
  if session.evaluate_script("window.angularReady === undefined") #variable is good to use
    #need to execute script to notify when angular finishes loading.
    session.execute_script <<-JS
    window.angularReady = false;
    var app = angular.element(document.querySelector('[ng-app]'));
    var injector = app.injector();
    injector.invoke(function($browser) {
    $browser.notifyWhenNoOutstandingRequests(function() {
      window.angularReady = true;
    });
    });
    JS
  end
  session.evaluate_script("window.angularReady")
end
# angular.element is an alias for angular's builtin jQuery if jQuery is not available.
# ['np-app'] is a key element specific to angular
# the injector kicks off the application and attempts to retrieve the object as requestor
# Invoke the method and supply the method arguments from the injector invoke(fn, [self], [locals])
# .notifyWhenNoOutstandingRequests - waits for the app to settle and then sets window.angularReady = true

def skip_advertisement(session)
  # session.save_page
  session.click_link("Skip this ad »") if session.has_content?("Skip this ad »")
end

def set_to_num_per_page(session) #TODO need to make this dynamic for selection! Currently defaults to 200.
  num_picker = session.find("#view-dropdown-top")
  num_picker.click
  # num_element = "showLimit=#{num};pageSet();showLimitDropdownTop=false"
  set_num = session.find('[ng-click="showLimit=200;pageSet();showLimitDropdownTop=false"]')
  set_num.click
end
#200 doesnt work
def switch_to_new_profile(session)
  open_windows = session.driver.browser.window_handles
  session.driver.browser.switch_to.window(open_windows.last)
  skip_advertisement(session)
end

def close_last_browser(session)
  session.driver.browser.close
  session.driver.browser.switch_to.window(session.driver.browser.window_handles[0])
end

def create_company_row(session)
  begin
    company_row = []
    businessName = session.all('h1.businessName')[0].text
    description = session.all('.description')[0].text
    session.has_text?("$")
    # binding.pry
    session.find('dl.RankTable').text.nil?
    rank = session.all('dl.RankTable :nth-child(1) dd')[0].text
    three_yr_growth = session.all('dl.RankTable :nth-child(2) dd')[0].text
    revenue_2013 = session.all('dl.RankTable :nth-child(3) dd')[0].text
    revenue_2010 = session.all('dl.RankTable :nth-child(4) dd')[0].text
    industry = session.all('dl.RankTable :nth-child(5) dd')[0].text

    # location table
    location = session.all('dl.LocationTable .row:nth-of-type(1) .dtdd:nth-of-type(1) dd')[0].text
    location_city = location.split(",")[0]
    location_state = location.split(",")[1]

    year_founded = session.all('dl.LocationTable .row:nth-of-type(1) .dtdd:nth-of-type(2) dd')[0].text
    employees = session.all('dl.LocationTable .row:nth-of-type(2) .dtdd:nth-of-type(1) dd')[0].text
    jobs_added_prev_3 = session.all('dl.LocationTable .row:nth-of-type(2) .dtdd:nth-of-type(2) dd')[0].text

    website = session.all('dl.LocationTable .row:nth-of-type(3) .dtdd:nth-of-type(1) dd')[0].text

    company_row.push(rank, businessName, description, three_yr_growth, revenue_2013, revenue_2010, industry, location_city, location_state, year_founded, employees, jobs_added_prev_3, website)
    return company_row
  rescue
    company_row = ["ERROR"]
  end
  # binding.pry
    # puts "Uh Oh 503"
    # session.reset_session!
end

def transfer_to_excel(header, companies)
  wb_name = "Inc5000 " + "#{Time.now.strftime("%m%d%Y %H%M")}" + ".xls"
  workbook = WriteExcel.new(wb_name)
  worksheet  = workbook.add_worksheet
  worksheet.write_row(0,0,header)
  index = 1
  companies.each do |company|
    worksheet.write_row(index, 0, company)
    index += 1
  end
  workbook.close
end

def start_cache(session, url)
  session.visit url
  skip_advertisement(session)
  make_angular_ready(session)
  set_to_num_per_page(session)
end

def go_to_page(session, num)
  pageInput = session.all('[ng-model="pageInput"]')[0]
  goButton = session.all('.goButton')[0]
  pageInput.set("#{num}")
  goButton.click
end


#TODO make session a class
header_row = ["Rank", "Name","Description","3-YR-Growth","2013 Revenue", "2010 Revenue", "Industry", "Location-City", "Location-State", "Year Founded", "Employees", "Jobs Added - Prev 3 Years", "Website"]
company_rows = []
session = Capybara::Session.new(:selenium)
main_url = "http://www.inc.com/inc5000/list/2014"
start_cache(session, main_url)

last_page = session.all('div.pages-container .page-box span').last.text.to_i
puts last_page

cur_page = 15
wb_name = "Inc5000 " + "#{Time.now.strftime("%m%d%Y %H%M")}" + ".xls"
workbook = WriteExcel.new(wb_name)
worksheet  = workbook.add_worksheet
worksheet.write_row(0,0,header_row)
excel_index = 1
while cur_page <= last_page

  # pageInput = session.all('[ng-model="pageInput"]')[0]
  # goButton = session.all('.goButton')[0]
  # pageInput.set("#{cur_page}")
  # goButton.click
  go_to_page(session, cur_page)

  # companies = session.all('[ng-repeat-start]')
  companies = 0
   #only the 4th child is visible from the search criteria.
  # binding.pry
  index = 0
  prior_index = -1
  num_per_page = 5000/last_page
  while companies < num_per_page
    begin
      names = session.all('.data-cell.bold a:nth-child(4)')
      names[index].click #click to open profile
      switch_to_new_profile(session) #current url is now profiles
      puts session.current_url
      unless session.current_url == "http://www.inc.com/inc5000/list/2014"
        # binding.pry
        company = create_company_row(session)
        company_rows << company
        puts company.to_s
        worksheet.write_row(excel_index, 0, company)
        excel_index += 1
        close_last_browser(session)
        index += 1
        prior_index += 1
        sleep rand(5..8)
      else
        prior_index += 1
        close_last_browser(session)
      end

      if prior_index == index
        start_cache(session, main_url)
        go_to_page(session, cur_page)
        names = session.all('.data-cell.bold a:nth-child(4)')
        names[index].click #click to open profile
        switch_to_new_profile(session) #current url is now profiles
        puts session.current_url
        company = create_company_row(session)
        puts company.to_s
        company_rows << company
        worksheet.write_row(excel_index, 0, company)
        excel_index += 1
        close_last_browser(session)
        index += 1
        sleep rand(2..4)
      end
      companies += 1
    rescue
      worksheet.write_row(excel_index, 0, ["error"])
      transfer_to_excel(header_row, company_rows)
      # retry
    end
  end
    puts cur_page.to_s
    cur_page += 1
    transfer_to_excel(header_row, company_rows)
end
workbook.close
puts company_rows
transfer_to_excel(header_row, company_rows)