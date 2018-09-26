Rails.application.routes.draw do
  root to: "products#index"

  resources :products
  resources :categories
  resources :groups

  post "products/increment" => "products#increment"
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
