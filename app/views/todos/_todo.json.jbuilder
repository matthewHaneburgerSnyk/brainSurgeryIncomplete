json.extract! todo, :id, :title, :notes, :code, :created_at, :updated_at
json.url todo_url(todo, format: :json)
