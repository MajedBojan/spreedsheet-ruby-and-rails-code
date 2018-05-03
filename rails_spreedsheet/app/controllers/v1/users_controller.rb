class V1::UsersController < ApplicationController

  def index
    users = User.all

    respond_to do |format|

      format.json { render json: {users: users } }

      format.xls  do
        task = Spreadsheet::Workbook.new
        sheet = task.create_worksheet

        rows_format = Spreadsheet::Format.new color: :purple,
        weight: :normal,
        size: 13,
        align: :center

        users.each.with_index(1) do |task, i|
          sheet.row(i).concat task.slice(:name, :email, :nationality_no, :birth_date).values

          sheet.row(i).height = 25
          sheet.column(i).width = 30
          sheet.row(i).default_format = rows_format
        end

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

        temp_file = StringIO.new
        task.write(temp_file)
        send_data(temp_file.string, :filename => "users.xls", :disposition => 'inline')
      end
    end

  end
end
