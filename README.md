### This Project created to explain spreadsheet gem in `Ruby` and `Ruby on Rails` as simple as i can

# First Let's start with Ruby
Let's go directly in the topic

First Create `.rb` file call it ruby spreadsheet or call it whatever name you like
in ruby there are no gems but we can require whatever library we want so let's add `require 'spreadsheet'` on top of the file we already created

will create new worksheet and assign it on Variable so we can reuse it

`book = Spreadsheet::Workbook.new`
`sheet1 = book.create_worksheet`
now we created one plain worksheet
we can use specific name for our file by adding this line `sheet1.name = 'Test spreadsheet'`

for row zero in our file will keep it for headers
`sheet1.row(0).concat %w{Name Country speacialist}`

here we go now we have done to much just we have to get our list we can get the list from anywhere but in my case i have amazing list for my teammate and some of the best developers in [Clickapps](http://www.clickapps.co/en/)

```row = sheet1.row(1)
row[0] = 'Majed Bojan'
row[1] =  'Yemen'
row[2] = 'Fullstack Developer'

row = sheet1.row(2)
row[0] = 'Ali Sheiba'
row[1] =  'Yemen'
row[2] = 'Fullstack Developer'

row = sheet1.row(3)
row[0] = 'Mohammed Balfaqi'
row[1] =  'Yemen'
row[2] = 'ROR Developer'

row = sheet1.row(4)
row[0] = 'Mohammed Basalah'
row[1] =  'Yemen'
row[2] = 'ROR Developer'

row = sheet1.row(5)
row[0] = 'Mohammed Aljefry'
row[1] =  'Yemen'
row[2] = 'Fullstack Developer'
```
here we go now we already done some amazing steps just we need to setup the format we want and it's very simple

will select 18 for the height of the cell, you can prepare your format as you like
`sheet1.row(0).height = 18`
then will add some optional format

`format = Spreadsheet::Format.new :color => :blue,
                                 :weight => :bold,
                                 :size => 18`
`sheet1.row(0).default_format = format`

`bold = Spreadsheet::Format.new :weight => :bold`

finally will choose the location we need to save our file in my case i will save it in the desktop  
book.write '/home/bojan/Desktop/spreedsheet.xls'

# Second Let's give try with Rails and exporting users from DB
For rails first we have to create rails project and will call it `rails_spreedsheet`
or you can call it what ever you want

let's setup rails app and bundling it
`rails new rails_spreedsheet --api --database=postgresql`
we have to install some dependencies as will need them, so please copy those gems into your gem file then run `bundle`

```
gem 'factory_bot_rails' # factory to gererate random records
gem 'spreadsheet'       # gem to convert into xls files
gem 'faker'             # to generate a random data
```
after you run `bundle` prepare `database.yml` and `secrets.yml` and ignore them then create examples for them

Now let's generate user model to and put some faker data
`rails g model user name:string email:string birth_date:datetime nationality_no:string`
now our migration file looks great so let's migrate it by runing `rails db:migrate`

We have user model let's get some data to export latter we will generate data using faker gem we can get number of users we want

copy this code to your seed or you can run it in `rails console` directly

```
100.times do
  User.create(
    name:           Faker::Name.name_with_middle,
    email:          Faker::Internet.email,
    nationality_no: Faker::Number.number(5),
    birth_date:     Faker::Date.birthday(18, 65)
  )
end
```
