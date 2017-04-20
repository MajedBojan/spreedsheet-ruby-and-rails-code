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
