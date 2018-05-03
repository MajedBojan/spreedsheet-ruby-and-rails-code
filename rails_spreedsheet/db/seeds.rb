100.times do
  User.create(
    name:           Faker::Name.name_with_middle,
    email:          Faker::Internet.email,
    nationality_no: Faker::Number.number(5),
    birth_date:     Faker::Date.birthday(18, 65)
  )
end
