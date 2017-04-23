require 'spreadsheet'


Spreadsheet.client_encoding = 'UTF-8'

# read from xls

# book = Spreadsheet.open '/home/clickapps/Desktop/excel-file.xls'
#
# # book.worksheets
#
# sheet1 = book.worksheet 0
#
# sheet1.each do |rows|
#
#
#   print   rows, "\n"
#   puts '=================================='
#
# end


# Write to xls

# will create new worksheet
book = Spreadsheet::Workbook.new

sheet1 = book.create_worksheet
sheet1.name = 'Test spreadsheet'


# or use This waye to to rename your spreadsheet
# sheet2 = book.create_worksheet :name => 'My Second Worksheet'


sheet1.row(0).concat %w{Name Country Acknowlegement}
# sheet1.row(1).concat %w{ Yukihiro_Matsumoto Japan Creator_of_Ruby'}
# sheet1.row(2).concat %w{ Yukihiro Matsumoto Japan Creator of Ruby'}

row = sheet1.row(1)
row[0] = 'Yukihiro Matsumoto'
row[1] =  'Japan'
row[2] = 'Creator of Ruby'

row = sheet1.row(2)
row[0] = 'MAjed'
row[1] =  'Yemen'
row[2] = 'developer of Ruby'


sheet1.row(0).height = 18

format = Spreadsheet::Format.new :color => :blue,
                                 :weight => :bold,
                                 :size => 18
sheet1.row(0).default_format = format

bold = Spreadsheet::Format.new :weight => :bold

book.write '/home/clickapps/Desktop/excel-file.xls'



#code for rails

first install spreadsheet gem

# gem to convert into xls files
gem 'spreadsheet', '~> 1.1', '>= 1.1.4'

then go to this pasth
config/initializers/mime_types.rb

and print this Mime::Type.register "application/xls", :xls


format.xls  do
  task = Spreadsheet::Workbook.new
  sheet = task.create_worksheet


  rows_format = Spreadsheet::Format.new color: :purple,
                                        weight: :normal,
                                        size: 13,
                                        align: :center

  @tasks.each.with_index(1) do |task, i|
    sheet.row(i).concat task.slice(:name, :description, :is_complete, :deadline, :employee_id).values
    sheet.row(i).height = 25
    sheet.column(i).width = 30
    sheet.row(i).default_format = rows_format

    # task.attributes.values
  end

  # save file
  # task.write '/home/clickapps/Desktop/test Spreadsheet.xls'
  head_format = Spreadsheet::Format.new color: :blue,
                                        weight: :bold,
                                        size:    14,
                                        pattern_bg_color:  :pattern_bg,
                                        pattern:  2,
                                        vertical_align: :middle,
                                        align:  :center


 sheet.row(0).concat %w{name description is_complete deadline employee_id}
 sheet.row(0).height = 25
 sheet.column(0).width = 30
 sheet.row(0).each.with_index { |c, i| sheet.row(0).set_format(i, head_format) }




  # bold = Spreadsheet::Format.new :weight => :bold


  temp_file = StringIO.new
  task.write temp_file
  temp_file.string.force_encoding('binary')
  send_data(temp_file, :filename => "tasks.xls", :disposition => 'inline')

end
end

# pagenation = @project.tasks.page(params[:page]).per(params[:per])
# render_data(data: ActiveModel::Serializer::CollectionSerializer.new(pagenation))
end

# GET: /v1/projects/:project_id/tasks/:task_id


######################################################################-
######################################################################-
######################################################################-
# praper it with service and spreedsheet files on employees controller


# service file

class EmployeesServices
  extend Enumerize
  # Attribiute to deffine columns attributes
  attr_accessor :header_columns
  # Enumerize header_columns to put attributes that not in model
  enumerize :header_columns, in: [:fullname, :birth_date, :joining_date, :salary,:designation ]

  def initialize(employees)
    @employees = employees
  end

  # Biuld report headers and records
  def biuld_report
    OpenStruct.new(
      header_columns: header_columns,
      records: records
    )
  end

  # Return employees attributes with add another employees
  def records
    @employees.map{ |employee| record(employee) }
  end

  def header_columns
    self.class.header_columns.options.map(&:first)
  end

  # Returns all record for specific employee
  def record employee
    {
      fullname: full_name(employee),
      birth_date: employee.birth_date,
      joining_date: employee.joining_date,
      salary:      employee.salary,
      designation: employee.designation,
    }
  end

  # Merge first_name and last_name
  def full_name(employee)
    "#{employee.first_name} #{employee.last_name}"
  end
end

######################################################################-
######################################################################-
######################################################################-
# spreadsheet file to fetch xls data and format

class EmployeesSheet

  def initialize(sheet, service_object)
    @service_object = service_object
    @sheet = sheet
  end

  # apply format for header and rows with different number type
  def apply_format
    rows_format = Spreadsheet::Format.new color: :purple,
                                          weight: :normal,
                                          size: 13,
                                          align: :center,
                                          pattern_bg_color:  :yellow,
                                          pattern:  2,
                                          :number_format => "#,##0.00 [$RS-407]"

    head_format = Spreadsheet::Format.new color: :blue,
                                          weight: :bold,
                                          size:    14,
                                          pattern_bg_color:  :pattern_bg,
                                          pattern:  2,
                                          vertical_align: :middle,
                                          align:  :center

    date_format = Spreadsheet::Format.new color: :purple,
                                          weight: :normal,
                                          size: 13,
                                          align: :center,
                                          pattern_bg_color:  :yellow,
                                          pattern:  2,
                                          :number_format => 'D-MMM-YYYY'

    @sheet.row(0).each.with_index do  |c, i| @sheet.row(0).set_format(i, head_format)
      @sheet.column(i).width = 30
      @sheet.row(i).height = 25
    end

    @service_object.records.each.with_index(1) do |employee, i|
      @sheet.row(i).each.with_index do  |c, x| @sheet.row(i).set_format(x, rows_format )
      @sheet.row(i).set_format(1,date_format )
      @sheet.row(i).set_format(2,date_format )
      @sheet.column(i).width = 30
      @sheet.row(i).height = 25
      end
    end
  end

  # Return header columns
  def biuld
    @sheet.row(0).concat @service_object.header_columns
    @service_object.records.each.with_index(1) do |employee, i|
      @sheet.row(i).concat employee.values
    end
  end
end

######################################################################-
######################################################################-
######################################################################-

# index what wii should be

# GET: v1/employees
def index

  service_object = EmployeesServices.new(@current_company.employees).biuld_report
  # filter = params.permit(:first_name_cont, :last_name_cont, :birth_date_cont, :joining_date_cont, :salary_eq)
  #
  # q = @current_company.employees.ransack(filter).result
  # employees = q.order("#{sort_column} #{sort_direction}")
  #
  # filter.present? ? render_data( data: {projects: employees}) :
  # render_data({ data: {employees: ActiveModel::Serializer::CollectionSerializer.new(employees)} })

  # @employees = @current_company.employees

  respond_to do |format|
    format.json {render_data(data: {employees: service_object.records})}

    format.pdf do
      html = ActionController::Base.new.render_to_string(
      template: 'v1/employees/index.pdf.erb',
      assigns:  { employees: service_object },
      orientation: 'Landscape',
      page_size: 'Letter',
      background: true
      )
      pdf = WickedPdf.new.pdf_from_string(html)
      send_data(pdf, :filename => "projects.pdf", :disposition => 'inline')
    end

    format.xls  do

      employees = Spreadsheet::Workbook.new
      sheet = employees.create_worksheet

      object_sheet = EmployeesSheet.new(sheet, service_object)

      object_sheet.biuld
      object_sheet.apply_format

      temp_file = StringIO.new
      employees.write temp_file
      temp_file.string.force_encoding('binary')
      send_data(temp_file, :filename => "employees.xls", :disposition => 'inline')
      end
  end
end
