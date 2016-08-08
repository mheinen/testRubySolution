Rails.application.routes.draw do

  root 'slides#index'
  post '/se' => 'slides#se'


end
