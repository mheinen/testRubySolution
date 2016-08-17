Rails.application.routes.draw do

  root 'slides#index'
  post '/se' => 'slides#se'
  get 'slides/download_pdf'
  post '/plugin' => 'slides#plugin'
  get '/videos' => 'slides#videos'


end
