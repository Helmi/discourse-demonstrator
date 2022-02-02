# frozen_string_literal: true

# name: discourse-demonstrator
# about: Send invites to new demonstrators
# version: 0.2.1
# authors: Pfaffman, Helmi
# url: https://github.com/helmi/discourse-demonstrator

#gem 'gimite-google-spreadsheet-ruby', '0.0.5', { require: false }
gem 'ruby-ole', '1.2.12', { require: false }
gem "spreadsheet", "1.2.0", { require: false }

enabled_site_setting :demonstrator_enabled
load File.expand_path('lib/demonstrator/demonstrator.rb', __dir__)

after_initialize do
  load File.expand_path('../app/jobs/regular/process_topic.rb', __FILE__)

  add_model_callback(Post, :after_create) do
    if self.topic.archetype != Archetype.private_message && self.topic.category.id == SiteSetting.demonstrator_category.to_i && self.post_number == 1
      Jobs::ProcessTopic.new.execute(topic_id: self.topic.id)
    end
  end
end
