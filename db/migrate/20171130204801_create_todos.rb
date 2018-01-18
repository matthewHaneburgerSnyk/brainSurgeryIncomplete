class CreateTodos < ActiveRecord::Migration[5.1]
  def change
    create_table :todos do |t|
      t.string :title
      t.text :notes
      t.text :code

      t.timestamps
    end
  end
end
