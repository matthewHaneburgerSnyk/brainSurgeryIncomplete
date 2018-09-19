require_relative 'boot'

require 'rails/all'

# Require the gems listed in Gemfile, including any gems
# you've limited to :test, :development, or :production.
Bundler.require(*Rails.groups)

module GraphingTool
  class Application < Rails::Application
    # Initialize configuration defaults for originally generated Rails version.
    config.load_defaults 5.1

    # Settings in config/environments/* take precedence over those specified here.
    # Application configuration should go into files in config/initializers
    # -- all .rb files in that directory are automatically loaded.
    
    #config.active_record.raise_in_transactional_callbacks = true
    config.generators do |g|
    g.test_framework :rspec
  end
  config.autoload_paths += %W(#{config.root}/lib)
  config.exceptions_app = routes
  end
end
