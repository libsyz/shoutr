class CreateShouts < ActiveRecord::Migration[5.0]
  def change
    create_table :shouts do |t|
      t.text :body, null: false
      t.references :user, foreign_key: true

      t.timestamps
    end
  end
end
