class CreatePptvideos < ActiveRecord::Migration
  def change
    create_table :pptvideos do |t|
      t.string :path
      t.text :title
      t.text :transition
      t.text :notes

      t.timestamps null: false
    end
  end
end
