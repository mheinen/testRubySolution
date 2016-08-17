class Pptvideo < ActiveRecord::Base
  serialize :title,Array
  serialize :transition,Array
  serialize :notes,Array
  validates_uniqueness_of :path
end
