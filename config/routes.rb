Rails.application.routes.draw do

  get 'shouts/create'

  get 'dashboards/show'

  constraints Clearance::Constraints::SignedIn.new do
  root to: 'dashboards#new'
  end

  root 'homes#show'
  resources :shouts, only: [:create]
  resources :passwords, controller: "clearance/passwords", only: [:create, :new]
  resource :session, only: [:create]

  resources :users, only: [:create] do
    resource :password,
      controller: "clearance/passwords",
      only: [:create, :edit, :update]
  end

  get "/sign_in" => "sessions#new", as: "sign_in"
  delete "/sign_out" => "sessions#destroy", as: "sign_out"
  get "/sign_up" => "users#new", as: "sign_up"
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
