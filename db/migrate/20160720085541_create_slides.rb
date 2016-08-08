class CreateSlides < ActiveRecord::Migration
  def change
    create_table :slides do |t|
      t.string :content
      t.string :title
      t.string :images

      t.timestamps null: false
    end
  end
end
