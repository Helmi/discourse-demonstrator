# frozen_string_literal: true

# name: discourse-demonstrator
# about: Send invites to new demonstrators
# version: 0.1.6
# authors: Pfaffman, Helmi
# url: https://github.com/helmi/discourse-demonstrator

#gem 'gimite-google-spreadsheet-ruby', '0.0.5', { require: false }
gem 'ruby-ole', '1.2.12', { require: false }
gem "spreadsheet", "1.2.0", { require: false }

enabled_site_setting :demonstrator_enabled
load File.expand_path('lib/demonstrator/demonstrator.rb', __dir__)

after_initialize do
  load File.expand_path('../app/jobs/regular/process_topic.rb', __FILE__)

  add_model_callback(Topic, :after_create) do
    # TODO: put in job rather than wait on post
    d_cat = Category.find(SiteSetting.demonstrator_category)
    Jobs.enqueue(:process_topic, topic_id: self.id) if d_cat.id == self.category.id
  end
end
