class DashboardsController < ApplicationController
  def new
    @shout = Shout.new
  end
end
